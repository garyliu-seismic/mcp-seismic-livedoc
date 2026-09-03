export interface ApiResult {
  status: number;
  body: unknown;
}

export interface LatestTokenData {
  formToken: string;
  writtenAt: number;
}

export interface SubmittedInputs {
  adHocInputs: Array<{ name: string; value: unknown }>;
  variableListData?: unknown;
  manualSelectContentInput?: unknown;
}

export interface ResultFileData {
  generatedLivedocId: string;
  status: string;
  downloadUrls: string[];
  downloads?: Array<{ url: string; format: string; fileName: string }>;
  templateName?: string;
  // The actual field values the user submitted in the App panel — NOT necessarily the same as
  // whatever prefill_livedoc_form_values suggested, since the user can edit before submitting.
  // Lets a later tool call (e.g. submit_ucb_workspace_generation) reuse the real submitted values
  // instead of the model having to recall/guess them from earlier turns.
  submittedInputs?: SubmittedInputs;
}

export interface PrefillData {
  scalars?: Record<string, unknown>;
  tables?: Record<string, Array<Record<string, unknown>>>;
  variableLists?: Record<string, {
    scalars?: Record<string, unknown>;
    tables?: Record<string, Array<Record<string, unknown>>>;
  }>;
}
