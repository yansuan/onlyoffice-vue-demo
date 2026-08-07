# Build frontend
FROM node:22-alpine AS build
WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci

COPY . .
# Vite values are placeholders; runtime /config.js (from -e) overrides them.
ENV VITE_DOCUMENT_SERVER_URL=http://127.0.0.1:19101
ENV VITE_CALLBACK_BASE_URL=http://127.0.0.1:19102
ENV VITE_ONLYOFFICE_JWT_SECRET=change-me
ENV VITE_DOCUMENT_PATH=/files/demo.docx
RUN npm run build-only

# Runtime: Express serves API + built SPA
FROM node:22-alpine
WORKDIR /app

ENV NODE_ENV=production
ENV PORT=4000
ENV STATIC_DIR=/app/dist

COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY --from=build /app/dist ./dist
COPY server ./server

EXPOSE 4000
CMD ["node", "server/callback-server.js"]
