import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  /* config options here */
  env: {
    USE_COGNITO_FEDERATION: process.env.USE_COGNITO_FEDERATION,
    MICROSOFT_ISSUER: process.env.MICROSOFT_ISSUER,
    COGNITO_REGION: process.env.COGNITO_REGION,
    COGNITO_CLIENT_ID: process.env.COGNITO_CLIENT_ID,
    COGNITO_CLIENT_SECRET: process.env.COGNITO_CLIENT_SECRET,
    COGNITO_USER_POOL_ID: process.env.COGNITO_USER_POOL_ID,
    COGNITO_EXTERNAL_PROVIDER: process.env.COGNITO_EXTERNAL_PROVIDER,
    AZURE_CLIENT_SECRET: process.env.AZURE_CLIENT_SECRET,
  },
  webpack: (config, { isServer }) => {
    // Handle Adaptive Cards library module resolution
    config.resolve.fallback = {
      ...config.resolve.fallback,
      fs: false,
      path: false,
    };

    // Handle ES modules in node_modules
    config.module.rules.push({
      test: /\.m?js$/,
      type: 'javascript/auto',
      resolve: {
        fullySpecified: false,
      },
    });

    return config;
  },
};

export default nextConfig;
