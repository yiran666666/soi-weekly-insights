FROM mirror.gcr.io/library/node:20-slim
WORKDIR /app
# Deps are installed on the host (Moloco network's Zscaler TLS interception
# breaks npm registry fetches inside the container). All deps are pure-JS with
# no native binaries, so copying host node_modules into linux/amd64 is safe.
COPY . .
RUN mkdir -p data uploads
EXPOSE 3456
ENV PORT=3456
CMD ["node", "server.js"]
