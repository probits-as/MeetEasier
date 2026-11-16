# Stage 1: Build Go binary and Templ views
FROM --platform=$BUILDPLATFORM golang:1.27-alpine AS builder
WORKDIR /src

COPY go.mod go.sum ./
RUN go mod download && go install github.com/a-h/templ/cmd/templ@$(go list -m github.com/a-h/templ | awk '{print $2}')

COPY . .
RUN templ generate
ARG TARGETOS=linux
ARG TARGETARCH=amd64
ARG VERSION=dev
ARG COMMIT=none
ARG DATE=unknown
RUN CGO_ENABLED=0 GOOS=${TARGETOS} GOARCH=${TARGETARCH} go build -trimpath \
    -ldflags="-s -w -X main.version=${VERSION} -X main.commit=${COMMIT} -X main.date=${DATE}" \
    -o /out/meeteasier ./cmd/meeteasier

# Stage 2: Minimal runtime image
FROM alpine:3.24
LABEL org.opencontainers.image.source="https://github.com/probits-as/MeetEasier" \
      org.opencontainers.image.description="Multi-tenant conference room display server built in Go" \
      org.opencontainers.image.licenses="GPL-3.0-only"
RUN apk upgrade --no-cache && apk add --no-cache ca-certificates tzdata \
    && addgroup -g 1001 -S meeteasier \
    && adduser -u 1001 -S -G meeteasier meeteasier

WORKDIR /app
COPY --from=builder /out/meeteasier /app/meeteasier
COPY static /app/static

RUN mkdir -p /app/data /app/bootstrap && chown -R meeteasier:meeteasier /app

USER meeteasier
EXPOSE 8080

ENV DB_PATH=/app/data/meeteasier.db
ENV PORT=8080
ENV MEETEASIER_DB_PATH=/app/data/meeteasier.db
ENV MEETEASIER_PORT=8080

ENTRYPOINT ["/app/meeteasier"]
CMD ["serve"]
