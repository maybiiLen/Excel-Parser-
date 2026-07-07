import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // Emit a self-contained server (.next/standalone/server.js) so the Docker
  // image can run `node server.js` without the full node_modules tree.
  output: "standalone",
  // Pin the workspace root to this project. Without this, a stray
  // package-lock.json in a parent directory makes Next infer the wrong root.
  turbopack: {
    root: __dirname,
  },
};

export default nextConfig;
