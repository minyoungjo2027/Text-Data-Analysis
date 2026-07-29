export type AnalysisSessionPayload = {
  created_at?: string;
  session_id?: string;
  student_name: string;
  student_id: string;
  comments: string[];
  vocabulary: string[];
  similarity_matrix: number[][];
  last_step: number;
};

const appsScriptUrl = process.env.NEXT_PUBLIC_GOOGLE_APPS_SCRIPT_URL?.trim() ?? "";

export async function saveAnalysisToGoogleSheets(payload: AnalysisSessionPayload) {
  if (!appsScriptUrl) {
    return {
      ok: false,
      error: "Google Apps Script 웹 앱 URL이 설정되지 않았습니다.",
    };
  }

  try {
    await fetch(appsScriptUrl, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
    });

    return { ok: true };
  } catch (error) {
    return {
      ok: false,
      error: error instanceof Error ? error.message : "Google Sheets 저장에 실패했습니다.",
    };
  }
}
