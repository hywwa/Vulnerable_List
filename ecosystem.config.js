module.exports = {
  apps: [
    {
      name: 'vulnerable-list-generator',
      script: './dist/server.js',
      cwd: './',
      instances: 1,
      autorestart: true,
      watch: false,
      max_memory_restart: '1G',
      env: {
        NODE_ENV: 'production'
      }
    }
  ]
};