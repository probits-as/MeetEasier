# Stage 1: Build Go binary and Templ views
FROM --platform=$BUILDPLATFORM golang:1.27-alpine AS builder
WORKDIR /src

RUN go install github.com/a-h/templ/cmd/templ@v0.3.977

COPY go.mod go.sum ./
RUN go mod download

COPY . .
RUN templ generate
ARG TARGETOS=linux
ARG TARGETARCH=amd64
RUN CGO_ENABLED=0 GOOS=${TARGETOS} GOARCH=${TARGETARCH} go build -trimpath -ldflags="-s -w" -o /out/meeteasier ./cmd/meeteasier

# Stage 2: Minimal runtime image
FROM alpine:3.24
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

ENTRYPOINT ["/app/meeteasier"]
CMD ["serve"]
