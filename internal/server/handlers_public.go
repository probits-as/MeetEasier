package server

import (
	"encoding/json"
	"errors"
	"log/slog"
	"net/http"

	"github.com/go-chi/chi/v5"

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
	if err := singleroom.SingleRoomPage(tenant, room, alias, found, status.Initialized, explicitTenant).Render(r.Context(), w); err != nil {
		slog.Error("Render single room page failed", "error", err)
	}
}

func (s *Server) handleSingleRoomPartial(w http.ResponseWriter, r *http.Request) {
	tenant, _, err := s.resolveTenant(r)
	if err != nil {
		http.Error(w, "Internal Server Error", http.StatusInternalServerError)
		return
	}

	alias := chi.URLParam(r, "alias")
	status := s.calMgr.GetStatus(tenant)
	room, found := s.calMgr.GetRoomByAlias(tenant, alias)

	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	if err := singleroom.SingleRoomContent(tenant, room, found, status.Initialized).Render(r.Context(), w); err != nil {
		slog.Error("Render single room partial failed", "error", err)
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

