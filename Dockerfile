FROM node:18-alpine

# Set working directory
WORKDIR /usr/src/app

# Install global dependencies
RUN npm install -g @microsoft/generator-sharepoint@1.20.0 gulp-cli@2.3.0 yo@4.3.1

# Create non-root user
RUN addgroup -g 1001 -S nodejs && \
    adduser -S spfx -u 1001

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm ci --silent && npm cache clean --force

# Copy source code
COPY . .

# Change ownership to nodejs user
RUN chown -R spfx:nodejs /usr/src/app
USER spfx

# Expose ports for dev server
EXPOSE 4321 35729

# Keep container running but don't auto-start gulp serve
CMD ["tail", "-f", "/dev/null"]