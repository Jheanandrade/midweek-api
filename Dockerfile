# Etapa de build da aplicação Go
FROM golang:1.24 AS builder

WORKDIR /app

COPY go.mod go.sum ./
RUN go mod download

COPY . ./
RUN CGO_ENABLED=0 GOOS=linux GOARCH=amd64 go build -o app

# Etapa final - runtime
FROM debian:bullseye-slim

WORKDIR /root/

# Variável para localizar o binário do LibreOffice
ENV LIBREOFFICE_PATH=libreoffice

# Instala apenas o necessário para rodar o LibreOffice headless
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    libreoffice \
    ca-certificates \
    fonts-dejavu-core \
    fonts-dejavu-extra \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Copia o binário da aplicação
COPY --from=builder /app/app .

# Expõe a porta HTTP
EXPOSE 8080

# Comando de execução
CMD ["./app"]