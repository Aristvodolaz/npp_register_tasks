// pm2 start ecosystem.config.js
// Переменные БД и пароли читаются приложением из .env в корне проекта.
module.exports = {
  apps: [
    {
      name: 'npp-register-v2',
      cwd: __dirname,
      script: 'server/dist/main.js',
      instances: 1,
      exec_mode: 'fork',
      autorestart: true,
      max_memory_restart: '500M',
      time: true,
      env: {
        NODE_ENV: 'production',
        PORT: 3043,
      },
      out_file: 'logs/out.log',
      error_file: 'logs/error.log',
      merge_logs: true,
    },
  ],
};
