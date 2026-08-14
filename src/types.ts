export interface ApiResult {
  status: number;
  body: unknown;
}

export interface LatestTokenData {
  formToken: string;
  writtenAt: number;
}

export interface ResultFileData {
  generatedLivedocId: string;
  status: string;
  downloadUrls: string[];
  downloads?: Array<{ url: string; format: string; fileName: string }>;
  templateName?: string;
}
