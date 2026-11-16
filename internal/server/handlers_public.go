package server

import (
	"context"
	"crypto/subtle"
	"encoding/json"
	"errors"
	"fmt"
	"log/slog"
	"net/http"
	"strconv"
	"strings"

	"github.com/go-chi/chi/v5"

	"github.com/probits-as/meeteasier/internal/model"
	"github.com/probits-as/meeteasier/internal/store"
	"github.com/probits-as/meeteasier/views/flightboard"
	"github.com/probits-as/meeteasier/views/singleroom"
)

func (s *Server) handleFlightboardPage(w http.ResponseWriter, r *http.Request) {
	tenant, explicitTenant, err := s.resolveTenant(r)
	if err != nil {
		slog.Error("Resolve tenant failed", "error", err)
		http.Error(w, "Internal Server Error", http.StatusInternalServerError)
		return
	}

	status := s.calMgr.GetStatus(tenant)
	activeFilter := r.URL.Query().Get("filter")

	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	if err := flightboard.FlightboardPage(tenant, status, activeFilter, explicitTenant).Render(r.Context(), w); err != nil {
		slog.Error("Render flightboard page failed", "error", err)
	}
}

func (s *Server) handleFlightboardPartial(w http.ResponseWriter, r *http.Request) {
	tenant, explicitTenant, err := s.resolveTenant(r)
	if err != nil {
		http.Error(w, "Internal Server Error", http.StatusInternalServerError)
		return
	}

	status := s.calMgr.GetStatus(tenant)
	activeFilter := r.URL.Query().Get("filter")

	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	if err := flightboard.FlightboardRows(tenant, status, activeFilter, explicitTenant).Render(r.Context(), w); err != nil {
		slog.Error("Render flightboard partial failed", "error", err)
	}
}

func (s *Server) handleSingleRoomPage(w http.ResponseWriter, r *http.Request) {
	tenant, explicitTenant, err := s.resolveTenant(r)
	if err != nil {
		http.Error(w, "Internal Server Error", http.StatusInternalServerError)
		return
	}

	alias := chi.URLParam(r, "alias")
	status := s.calMgr.GetStatus(tenant)
	room, found := s.calMgr.GetRoomByAlias(tenant, alias)

	if !found && status.Initialized {
		w.WriteHeader(http.StatusNotFound)
	}
	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	if err := singleroom.SingleRoomPage(tenant, room, alias, found, status.Initialized, explicitTenant, "", "").Render(r.Context(), w); err != nil {
		slog.Error("Render single room page failed", "error", err)
	}
}

func (s *Server) handleSingleRoomPartial(w http.ResponseWriter, r *http.Request) {
	tenant, explicitTenant, err := s.resolveTenant(r)
	if err != nil {
		http.Error(w, "Internal Server Error", http.StatusInternalServerError)
		return
	}

	alias := chi.URLParam(r, "alias")
	status := s.calMgr.GetStatus(tenant)
	room, found := s.calMgr.GetRoomByAlias(tenant, alias)

	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	if err := singleroom.SingleRoomContent(tenant, room, found, status.Initialized, explicitTenant, "", "").Render(r.Context(), w); err != nil {
		slog.Error("Render single room partial failed", "error", err)
	}
}

func (s *Server) handleSingleRoomBook(w http.ResponseWriter, r *http.Request) {
	tenant, explicitTenant, err := s.resolveTenant(r)
	if err != nil {
		http.Error(w, "Internal Server Error", http.StatusInternalServerError)
		return
	}

	alias := chi.URLParam(r, "alias")
	if err := r.ParseForm(); err != nil {
		room, found := s.calMgr.GetRoomByAlias(tenant, alias)
		s.renderSingleRoomBookResponse(w, r, tenant, room, alias, found, explicitTenant, "", "Ugyldig forespørsel", http.StatusBadRequest)
		return
	}

	action := strings.TrimSpace(r.FormValue("action"))
	durationMinutes, _ := strconv.Atoi(strings.TrimSpace(r.FormValue("duration_minutes")))
	if choice := strings.TrimSpace(r.FormValue("booking_choice")); choice != "" {
		if parts := strings.SplitN(choice, ":", 2); len(parts) == 2 {
			action = strings.TrimSpace(parts[0])
			durationMinutes, _ = strconv.Atoi(strings.TrimSpace(parts[1]))
		}
	}

	req := model.BookingRequest{
		RoomAlias:       alias,
		Action:          action,
		DurationMinutes: durationMinutes,
		Subject:         strings.TrimSpace(r.FormValue("subject")),
		PIN:             strings.TrimSpace(r.FormValue("pin")),
	}

	room, statusCode, execErr := s.validateAndExecuteBooking(r.Context(), tenant, alias, req)
	if execErr != nil {
		_, found := s.calMgr.GetRoomByAlias(tenant, alias)
		s.renderSingleRoomBookResponse(w, r, tenant, room, alias, found || room.RoomAlias != "", explicitTenant, "", execErr.Error(), statusCode)
		return
	}

	flash := bookingSuccessMessage(req)
	s.renderSingleRoomBookResponse(w, r, tenant, room, alias, true, explicitTenant, flash, "", http.StatusOK)
}

func (s *Server) renderSingleRoomBookResponse(w http.ResponseWriter, r *http.Request, tenant model.Tenant, room model.Room, alias string, found, explicitTenant bool, flashMsg, errMsg string, statusCode int) {
	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	isHTMX := strings.EqualFold(r.Header.Get("HX-Request"), "true")
	if isHTMX {
		// Return 200 for HTMX partial swaps so error banners are rendered inline on the touch panel.
		w.WriteHeader(http.StatusOK)
		if err := singleroom.SingleRoomContent(tenant, room, found, true, explicitTenant, flashMsg, errMsg).Render(r.Context(), w); err != nil {
			slog.Error("Render single room booking response failed", "error", err)
		}
		return
	}

	if statusCode >= 400 {
		w.WriteHeader(statusCode)
	}
	if err := singleroom.SingleRoomPage(tenant, room, alias, found, true, explicitTenant, flashMsg, errMsg).Render(r.Context(), w); err != nil {
		slog.Error("Render single room booking page failed", "error", err)
	}
}

func (s *Server) handleAPIRoomBooking(w http.ResponseWriter, r *http.Request) {
	tenant, _, err := s.resolveTenant(r)
	if err != nil {
		writeJSON(w, http.StatusInternalServerError, map[string]string{"error": "Failed to resolve tenant"})
		return
	}

	r.Body = http.MaxBytesReader(w, r.Body, 1<<20)
	var req model.BookingRequest
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		writeJSON(w, http.StatusBadRequest, map[string]string{"error": "Invalid JSON: " + err.Error()})
		return
	}
	if strings.TrimSpace(req.RoomAlias) == "" {
		writeJSON(w, http.StatusBadRequest, map[string]string{"error": "room_alias is required"})
		return
	}

	room, statusCode, execErr := s.validateAndExecuteBooking(r.Context(), tenant, req.RoomAlias, req)
	if execErr != nil {
		writeJSON(w, statusCode, map[string]string{"error": execErr.Error()})
		return
	}

	writeJSON(w, http.StatusOK, map[string]any{
		"status":  "ok",
		"message": bookingSuccessMessage(req),
		"room":    room,
	})
}

func (s *Server) validateAndExecuteBooking(ctx context.Context, tenant model.Tenant, alias string, req model.BookingRequest) (model.Room, int, error) {
	if !tenant.BookingEnabled {
		room, _ := s.calMgr.GetRoomByAlias(tenant, alias)
		return room, http.StatusForbidden, errors.New("Rombooking er deaktivert for denne tenanten")
	}

	if tenant.HasBookingPIN {
		submitted := strings.TrimSpace(req.PIN)
		if submitted == "" || subtle.ConstantTimeCompare([]byte(submitted), []byte(tenant.BookingPIN)) != 1 {
			room, _ := s.calMgr.GetRoomByAlias(tenant, alias)
			return room, http.StatusUnauthorized, errors.New("Feil PIN-kode / passord")
		}
	}

	room, err := s.calMgr.BookRoom(ctx, tenant, alias, req)
	if err != nil {
		slog.Warn("Room booking failed", "tenant", tenant.Slug, "room", alias, "action", req.Action, "error", err)
		return room, http.StatusBadRequest, err
	}

	slog.Info("Room booking executed", "tenant", tenant.Slug, "room", alias, "action", req.Action, "duration_minutes", req.DurationMinutes)
	return room, http.StatusOK, nil
}

func bookingSuccessMessage(req model.BookingRequest) string {
	switch req.Action {
	case model.BookingActionBookNow:
		return fmt.Sprintf("Rommet er booket i %d minutter", req.DurationMinutes)
	case model.BookingActionBookAfter:
		return fmt.Sprintf("Rommet er booket i %d min etter pågående møte", req.DurationMinutes)
	case model.BookingActionExtend:
		return fmt.Sprintf("Møtet er forlenget med %d minutter", req.DurationMinutes)
	case model.BookingActionEndNow:
		return "Møtet er avsluttet og rommet er ledig"
	default:
		return "Booking oppdatert"
	}
}

func (s *Server) handleTenantAsset(w http.ResponseWriter, r *http.Request) {
	slug := chi.URLParam(r, "slug")
	key := chi.URLParam(r, "key")

	asset, err := s.store.GetAsset(r.Context(), slug, key)
	if errors.Is(err, store.ErrNotFound) {
		http.NotFound(w, r)
		return
	}
	if err != nil {
		http.Error(w, "Internal Server Error", http.StatusInternalServerError)
		return
	}

	w.Header().Set("Content-Type", asset.ContentType)
	w.Header().Set("Cache-Control", "public, max-age=3600")
	_, _ = w.Write(asset.Data)
}

func (s *Server) handleAPIRooms(w http.ResponseWriter, r *http.Request) {
	tenant, _, err := s.resolveTenant(r)
	if err != nil {
		writeJSON(w, http.StatusInternalServerError, map[string]string{"error": "Failed to resolve tenant"})
		return
	}
	status := s.calMgr.GetStatus(tenant)
	if status.LastError != "" && len(status.Rooms) == 0 {
		writeJSON(w, http.StatusBadGateway, map[string]string{"error": status.LastError})
		return
	}
	writeJSON(w, http.StatusOK, status.Rooms)
}

func (s *Server) handleAPIRoomlists(w http.ResponseWriter, r *http.Request) {
	tenant, _, err := s.resolveTenant(r)
	if err != nil {
		writeJSON(w, http.StatusInternalServerError, map[string]string{"error": "Failed to resolve tenant"})
		return
	}
	status := s.calMgr.GetStatus(tenant)
	if status.LastError != "" && len(status.Roomlists) == 0 {
		writeJSON(w, http.StatusBadGateway, map[string]string{"error": status.LastError})
		return
	}
	writeJSON(w, http.StatusOK, status.Roomlists)
}

func (s *Server) handleAPIHeartbeat(w http.ResponseWriter, r *http.Request) {
	writeJSON(w, http.StatusOK, map[string]string{"status": "OK"})
}

func writeJSON(w http.ResponseWriter, statusCode int, payload any) {
	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	w.WriteHeader(statusCode)
	_ = json.NewEncoder(w).Encode(payload)
}

