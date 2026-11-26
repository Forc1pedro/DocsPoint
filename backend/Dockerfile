# =================================================================
# SharePoint Clone - Backend API Multi-stage Dockerfile
# =================================================================

# Stage 1: Build & Dependencies
FROM node:20-alpine AS builder

WORKDIR /app

# Copy package configurations
COPY package*.json tsconfig*.json ./

# Install all dependencies (including devDependencies for build)
RUN npm install

# Copy application source
COPY src ./src

# Compile TypeScript
RUN npm run build || (mkdir -p dist && npx tsc)

# =================================================================
# Stage 2: Production Runtime
# =================================================================
FROM node:20-alpine AS runner

WORKDIR /app

ENV NODE_ENV=production
ENV PORT=5000

# Install curl for container healthcheck
RUN apk add --no-cache curl

# Copy production dependencies only
COPY package*.json ./
RUN npm install --omit=dev

# Copy compiled JavaScript from builder
COPY --from=builder /app/dist ./dist

# Run as non-root user for enhanced security
USER node

EXPOSE 5000

HEALTHCHECK --interval=15s --timeout=5s --retries=3 \
  CMD curl -f http://localhost:5000/health || exit 1

CMD ["node", "dist/server.js"]
