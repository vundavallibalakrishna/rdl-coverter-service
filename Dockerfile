FROM node:22-bookworm-slim AS dependencies
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --omit=dev

FROM node:22-bookworm-slim
RUN apt-get update \
  && apt-get install -y --no-install-recommends poppler-utils fontconfig \
  && rm -rf /var/lib/apt/lists/*
WORKDIR /app
ENV NODE_ENV=production \
    PORT=7070 \
    HOST=0.0.0.0 \
    RDL_TEMP_ROOT=/tmp/rdl-converter \
    RDL_FONT_DIR=/app/fonts \
    RDL_STRICT_FONTS=true \
    RDL_PDFTOPPM_PATH=pdftoppm
COPY --from=dependencies /app/node_modules ./node_modules
COPY package.json ./
COPY src ./src
RUN mkdir -p /tmp/rdl-converter /app/fonts \
  && chmod 700 /tmp/rdl-converter \
  && chown -R node:node /tmp/rdl-converter /app
USER node
EXPOSE 7070
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
  CMD node -e "fetch('http://127.0.0.1:7070/healthz').then(r=>{if(!r.ok)process.exit(1)}).catch(()=>process.exit(1))"
CMD ["node", "src/server.js"]
