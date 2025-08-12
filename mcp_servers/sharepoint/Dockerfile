FROM node:22-alpine AS builder

WORKDIR /app

COPY mcp_servers/sharepoint/package.json mcp_servers/sharepoint/package-lock.json ./

RUN npm install --ignore-scripts

COPY mcp_servers/sharepoint .

RUN npm run build

FROM node:22-slim AS release

WORKDIR /app

COPY --from=builder /app/dist /app/dist
COPY --from=builder /app/package.json /app/package.json
COPY --from=builder /app/package-lock.json /app/package-lock.json

RUN npm ci --omit=dev --ignore-scripts

EXPOSE 5000

ENTRYPOINT ["node", "dist/index.js"]