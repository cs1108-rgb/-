// Cross-platform production entry point (Windows cmd does not support VAR=value).
process.env.NODE_ENV='production';
process.env.MEASURE_HOST ||= '127.0.0.1';
await import('./server.js');
