export default {
  build: {
    rollupOptions: {
      external: [
        "/npm/@azure/msal-browser@4.30.0/dist/telemetry/BrowserPerformanceMeasurement.mjs/+esm",
      ],
    },
  },
};
