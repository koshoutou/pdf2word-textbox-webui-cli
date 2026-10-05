import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  output: "standalone",
  reactStrictMode: false,
  // 允许预览面板跨域访问
  allowedDevOrigins: ['*.space-z.ai'],
};

export default nextConfig;
