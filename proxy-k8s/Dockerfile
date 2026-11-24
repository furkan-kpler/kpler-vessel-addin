# Pin the base image by digest before production use (see README).
FROM node:22-alpine

WORKDIR /app

ENV NODE_ENV=production \
    PORT=8080

# No npm dependencies: copy only the server and its manifest.
COPY --chown=node:node package.json server.js ./

# Run as the unprivileged "node" user that ships with the image.
USER node

EXPOSE 8080

CMD ["node", "server.js"]
