'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

// Force HTTP configuration
build.serve.taskConfig.https = false;
build.serve.taskConfig.port = 4321;
build.serve.taskConfig.hostname = '0.0.0.0';

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);
  result.set('serve', result.get('serve-deprecated'));
  return result;
};

build.initialize(require('gulp'));

// Dockerfile instructions
// FROM node:18-alpine
//
// # Install required packages
// RUN apk add --no-cache openssl ca-certificates curl git
//
// # Set working directory
// WORKDIR /usr/src/app
//
// # Install global dependencies with specific versions for compatibility
// RUN npm install -g @microsoft/generator-sharepoint@1.20.0 gulp-cli@2.3.0 yo@4.3.1
//
// # Create non-root user
// RUN addgroup -g 1001 -S nodejs && \
//     adduser -S spfx -u 1001
//
// # Create directories and set permissions
// RUN mkdir -p /home/spfx/.rushstack /home/spfx/.config && \
//     chown -R spfx:nodejs /home/spfx
//
// # Copy package files
// COPY package*.json ./
//
// # Install dependencies
// RUN npm ci --silent && npm cache clean --force
//
// # Copy source code
// COPY . .
