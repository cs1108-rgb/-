// The cloud development server must accept connections from its preview proxy.
process.env.MEASURE_HOST ||= '0.0.0.0';
await import('./server.js');
