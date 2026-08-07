# Build frontend
FROM node:22-alpine AS build
WORKDIR /app

# 默认阿里云 npm 源；覆盖示例: --build-arg NPM_REGISTRY=https://registry.npmjs.org/
ARG NPM_REGISTRY=https://registry.npmmirror.com
RUN npm config set registry "$NPM_REGISTRY"

COPY package.json package-lock.json .npmrc ./
RUN npm ci

COPY . .
# Placeholders only; runtime /config.js (docker -e) overrides them.
ENV VITE_DOCUMENT_SERVER_URL=http://127.0.0.1:19101
ENV VITE_CALLBACK_BASE_URL=http://127.0.0.1:19102
ENV VITE_DOCUMENT_PATH=/files/demo.docx
RUN npm run build-only

# Runtime: Express serves API + built SPA
FROM node:22-alpine
WORKDIR /app

ARG NPM_REGISTRY=https://registry.npmmirror.com
RUN npm config set registry "$NPM_REGISTRY"

ENV NODE_ENV=production
ENV PORT=4000
ENV STATIC_DIR=/app/dist

COPY package.json package-lock.json .npmrc ./
RUN npm ci --omit=dev

COPY --from=build /app/dist ./dist
COPY server ./server

EXPOSE 4000
CMD ["node", "server/callback-server.js"]
