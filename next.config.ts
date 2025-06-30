import type { NextConfig } from "next";

// Log environment variables during build
console.log("ENV VARS AT BUILD TIME:");
for (const [key, value] of Object.entries(process.env)) {
  console.log(`${key}=${value}`);
}

const nextConfig: NextConfig = {
  /* config options here */
};

export default nextConfig;
