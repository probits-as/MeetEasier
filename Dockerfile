# Stage 1: Build React UI
FROM node:18-alpine AS ui-builder
WORKDIR /app/ui-react
COPY ui-react/package*.json ./
RUN npm install
COPY ui-react/ ./
RUN npm run build

# Stage 2: Run Express server
FROM node:18-alpine AS server
WORKDIR /app

COPY package*.json ./
RUN npm install --ignore-scripts

COPY . .

# Copy built UI from stage 1
COPY --from=ui-builder /app/ui-react/build ./ui-react/build

EXPOSE 8080

CMD ["node", "server.js"]
