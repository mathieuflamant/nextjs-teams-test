import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  /* config options here */
  env: {
    USE_COGNITO_FEDERATION: process.env.USE_COGNITO_FEDERATION,
    MICROSOFT_ISSUER: process.env.MICROSOFT_ISSUER,
    COGNITO_TOKEN_ENDPOINT: process.env.COGNITO_TOKEN_ENDPOINT,
    COGNITO_CLIENT_ID: process.env.COGNITO_CLIENT_ID,
    COGNITO_CLIENT_SECRET: process.env.COGNITO_CLIENT_SECRET,
    COGNITO_USER_POOL_ID: process.env.COGNITO_USER_POOL_ID,
    COGNITO_REGION: process.env.COGNITO_REGION,
    AZURE_CLIENT_SECRET: process.env.AZURE_CLIENT_SECRET,
  },
};

export default nextConfig;
