FROM node:22-bookworm-slim

ENV NODE_ENV=production
ENV PORT=4000

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-impress fonts-liberation fontconfig python3 make g++ \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY package*.json ./
RUN npm install --no-audit --no-fund
COPY . .
RUN npm run build:web

EXPOSE 4000
CMD ["npm", "run", "start:server"]
