"use client";

import { useEffect, useMemo, useState } from "react";
import { saveAnalysisToGoogleSheets } from "../lib/google-sheets";

declare global {
  interface Window {
    XLSX?: any;
    html2pdf?: any;
  }
}

const stopwords = new Set([
  "은", "는", "이", "가", "을", "를", "에", "의", "와", "과", "도", "으로", "로", "에서", "에게", "한테", "께", "부터", "까지", "보다", "처럼", "만큼", "조차", "마저", "밖에", "뿐", "마다", "라고", "이라", "하고", "하며", "그리고", "하지만", "그러나", "그래서", "또한", "정말", "너무", "매우", "아주", "잘", "더", "가장", "있다", "있는", "있고", "하다", "하는", "한", "좋다", "좋은", "것", "수", "등", "및", "또", "안", "못", "않다", "어떤", "이런", "저런", "그런", "모든", "각", "여러", "다른", "한", "두"
]);
const attachedParticles = ["으로", "에서", "에게", "한테", "까지", "부터", "처럼", "만큼", "보다", "이라", "라고", "하고", "은", "는", "이", "가", "을", "를", "에", "의", "와", "과", "도", "로", "만"].sort((a, b) => b.length - a.length);
const initial = [
  "달콤, 상큼한 성심당 망고시루 진짜 최고예요! 망고시루 살 때 직원분도 너무 친절하셔서 감동했어요. 다음에도 상큼한 망고시루 사러 친절한 성심당 또 갈게요.",
  "망고시루 소문 듣고 성심당 다녀왔는데 망고시루 맛은 물론이고 친절한 서비스 덕분에 기분 좋았어요. 같이 산 순수롤도 대박이네요!",
  "성심당 순수롤은 정말 촉촉하고 순수롤 크림이 최고예요. 순수롤 생각나서 성심당 매일 가고 싶어요. 직원분도 아주 친절했습니다.",
  "성심당 순수롤이랑 망고시루 둘 다 사서 먹었는데 순수롤이 생각보다 훨씬 부드럽고 순수롤 달콤해요. 순수롤 강력 추천합니다!",
  "성심당 친절한 응대 덕분에 구매할 때 기분이 좋았어요. 역시 친절이 기본이 되는 매장이라 친절함에 반해서 성심당 자주 찾게 되네요.",
];
const legacyInitialComment = "성심당 망고시루 진짜 최고예요! 망고시루 살 때 직원분도 너무 친절하셔서 감동했어요. 다음에도 망고시루 사러 친절한 성심당 또 갈게요.";
const clusterPalette = [
  { background: "#fff2d9", border: "#d49a4d" },
  { background: "#f8e0d3", border: "#b85c3d" },
  { background: "#ece2dc", border: "#6d4c41" },
];
const STORAGE_KEY = "textlab-analysis-session";
type TfidfFilter = "all" | "common" | "top5" | "top8";

function createSessionId() {
  return typeof crypto !== "undefined" && "randomUUID" in crypto
    ? crypto.randomUUID()
    : `session-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

function tokens(text: string) {
  return text
    .replace(/[^가-힣a-zA-Z0-9\s]/g, " ")
    .split(/\s+/)
    .filter((word) => word.length > 1)
    .map((word) => word.toLowerCase())
    .flatMap((word) => {
      // 한 글자 명사도 조사와 분리합니다. 예: "빵은" → "빵" + "은"
      const particle = attachedParticles.find((suffix) => word.endsWith(suffix) && word.length - suffix.length >= 1);
      return particle ? [word.slice(0, -particle.length), particle] : [word];
    });
}

// 교육용 분석에서는 자주 쓰이는 활용형을 대표 단어(표제어)로 묶습니다.
// 예: 친절하셔서·친절했습니다·친절한·친절함 → 친절
function normalizeWord(word: string) {
  const exact: Record<string, string> = {
    친절하셔서: "친절",
    친절했습니다: "친절",
    친절합니다: "친절",
    친절했어요: "친절",
    친절해서: "친절",
    친절하다: "친절",
    친절해요: "친절",
    친절하게: "친절",
    친절한: "친절",
    친절함: "친절",
    상큼한: "상큼",
    순수롤이랑: "순수롤",
  };
  return exact[word] || word;
}

export default function Home() {
  const [comments, setComments] = useState(initial);
  const [studentName, setStudentName] = useState("");
  const [studentId, setStudentId] = useState("");
  const [step, setStep] = useState(1);
  const [pdfMode, setPdfMode] = useState(false);
  const [saveState, setSaveState] = useState<"idle" | "saving" | "saved" | "error">("idle");
  const [saveError, setSaveError] = useState("");
  const [savedAt, setSavedAt] = useState("");
  const [sessionId, setSessionId] = useState("");
  const [motionEnabled, setMotionEnabled] = useState(true);
  const [selectedCommon, setSelectedCommon] = useState("");
  const [tfidfFilter, setTfidfFilter] = useState<TfidfFilter>("all");

  useEffect(() => {
    const stored = window.localStorage.getItem(STORAGE_KEY);
    if (!stored) return;
    try {
      const session = JSON.parse(stored) as { studentName?: string; studentId?: string; comments?: string[]; savedAt?: string; sessionId?: string };
      if (session.studentName) setStudentName(session.studentName);
      if (session.studentId) setStudentId(session.studentId);
      if (session.comments?.length) {
        const comments = session.comments.map((comment, index) => index === 0 && comment === legacyInitialComment ? initial[0] : comment);
        setComments(comments);
        if (comments[0] !== session.comments[0]) {
          window.localStorage.setItem(STORAGE_KEY, JSON.stringify({ ...session, comments }));
        }
      }
      if (session.savedAt) setSavedAt(session.savedAt);
      setSessionId(session.sessionId || createSessionId());
    } catch {
      window.localStorage.removeItem(STORAGE_KEY);
      setSessionId(createSessionId());
    }
    if (!stored) setSessionId(createSessionId());
  }, []);
  const data = useMemo(() => {
    const raw = comments.map(tokens);
    const cleaned = raw.map((words) => words.filter((word) => !stopwords.has(word)).map(normalizeWord));
    const vocabulary = [...new Set(cleaned.flat())];
    const df = Object.fromEntries(vocabulary.map((word) => [word, cleaned.filter((doc) => doc.includes(word)).length]));
    const idf = Object.fromEntries(vocabulary.map((word) => [word, Math.log((comments.length + 1) / ((df[word] || 0) + 1)) + 1]));
    const vectors = cleaned.map((doc) => vocabulary.map((word) => doc.filter((w) => w === word).length * idf[word]));
    const cosine = (a: number[], b: number[]) => {
      const denom = Math.sqrt(a.reduce((s, v) => s + v * v, 0)) * Math.sqrt(b.reduce((s, v) => s + v * v, 0));
      return denom ? a.reduce((s, v, i) => s + v * b[i], 0) / denom : 0;
    };
    const tfidfRows = vocabulary.map((word, index) => {
      const values = vectors.map((vector) => vector[index]);
      return { word, values, score: Math.max(0, ...values), documentCount: df[word] ?? 0 };
    }).sort((a, b) => b.score - a.score);
    const embeddingRows = tfidfRows.slice(0, 5);
    const embeddingIndexes = embeddingRows.map((item) => vocabulary.indexOf(item.word));
    const embeddingVectors = vectors.map((vector) => embeddingIndexes.map((index) => vector[index] ?? 0));
    return { raw, cleaned, vocabulary, df, idf, vectors, tfidfRows, embeddingRows, embeddingVectors, matrix: embeddingVectors.map((a) => embeddingVectors.map((b) => cosine(a, b))) };
  }, [comments]);
  const visibleTfidfRows = data.tfidfRows.filter((item, index) => {
    if (tfidfFilter === "common") return item.documentCount >= 2;
    if (tfidfFilter === "top5") return index < 5;
    if (tfidfFilter === "top8") return index < 8;
    return true;
  });
  const idfRows = [...data.vocabulary].sort((a, b) => data.idf[a] - data.idf[b] || a.localeCompare(b, "ko")).slice(0, 12);
  const clusterGroups = useMemo(() => {
    const candidates = (data.tfidfRows.filter((item) => item.documentCount >= 2).length ? data.tfidfRows.filter((item) => item.documentCount >= 2) : data.tfidfRows).slice(0, 3);
    const assigned = new Set<number>();
    const groups = candidates.map((item, index) => {
      const documentIndexes = data.cleaned.map((doc, docIndex) => doc.includes(item.word) ? docIndex : -1).filter((docIndex) => docIndex >= 0 && !assigned.has(docIndex));
      documentIndexes.forEach((docIndex) => assigned.add(docIndex));
      const keywords = data.tfidfRows.filter((row) => documentIndexes.some((docIndex) => row.values[docIndex] > 0)).slice(0, 4).map((row) => row.word);
      return { label: `${String.fromCharCode(65 + index)} 그룹`, word: item.word, documentIndexes, keywords };
    }).filter((group) => group.documentIndexes.length > 0);
    const remaining = comments.map((_, index) => index).filter((index) => !assigned.has(index));
    if (remaining.length) groups.push({ label: "기타 그룹", word: "기타", documentIndexes: remaining, keywords: data.tfidfRows.filter((row) => remaining.some((docIndex) => row.values[docIndex] > 0)).slice(0, 4).map((row) => row.word) });
    return groups;
  }, [comments, data]);
  const commentGroupIndexes = useMemo(() => comments.map((_, commentIndex) => clusterGroups.findIndex((group) => group.documentIndexes.includes(commentIndex))), [comments, clusterGroups]);
  const largestClusterIndex = useMemo(() => clusterGroups.reduce((largest, group, index) => group.documentIndexes.length > (clusterGroups[largest]?.documentIndexes.length ?? 0) ? index : largest, -1), [clusterGroups]);
  const bestPair = useMemo(() => data.matrix.flatMap((row, i) => row.map((value, j) => ({ value, i, j }))).filter((pair) => pair.i !== pair.j).sort((a, b) => b.value - a.value)[0] ?? null, [data.matrix]);

  const update = (i: number, value: string) => setComments((prev) => prev.map((item, index) => (index === i ? value : item)));
  const saveAnalysis = async () => {
    if (!studentName.trim() || !studentId.trim()) {
      setSaveError("학생 이름과 학번을 먼저 입력해 주세요.");
      setSaveState("error");
      return;
    }
    setSaveState("saving");
    setSaveError("");
    const timestamp = new Date().toISOString();
    const currentSessionId = sessionId || createSessionId();
    setSavedAt(timestamp);
    setSessionId(currentSessionId);
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify({
      studentName: studentName.trim(),
      studentId: studentId.trim(),
      comments,
      savedAt: timestamp,
      sessionId: currentSessionId,
    }));
    const result = await saveAnalysisToGoogleSheets({
      created_at: timestamp,
      session_id: currentSessionId,
      student_name: studentName.trim(),
      student_id: studentId.trim(),
      comments,
      vocabulary: data.vocabulary,
      similarity_matrix: data.matrix,
      last_step: step,
    });
    if (!result.ok) {
      setSaveError(result.error ?? "Google Sheets 저장에 실패했습니다.");
      setSaveState("error");
    } else {
      setSaveState("saved");
    }
  };
  const downloadExcel = () => {
    if (!studentName.trim() || !studentId.trim()) {
      setSaveError("학생 이름과 학번을 먼저 입력해 주세요.");
      setSaveState("error");
      return;
    }
    if (!window.XLSX) {
      setSaveError("엑셀 라이브러리를 불러오지 못했습니다. 잠시 후 다시 시도해 주세요.");
      setSaveState("error");
      return;
    }
    const headers = ["created_at", "session_id", "student_name", "student_id", "last_step", "comments", "vocabulary", "similarity_matrix"];
    const row = [
      savedAt || new Date().toISOString(),
      sessionId || createSessionId(),
      studentName.trim(),
      studentId.trim(),
      step,
      JSON.stringify(comments),
      JSON.stringify(data.vocabulary),
      JSON.stringify(data.matrix),
    ];
    const worksheet = window.XLSX.utils.aoa_to_sheet([headers, row]);
    const workbook = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(workbook, worksheet, "텍스트 데이터");
    const safeId = studentId.trim().replace(/[\\/:*?"<>|]/g, "_");
    const safeName = studentName.trim().replace(/[\\/:*?"<>|]/g, "_");
    window.XLSX.writeFile(workbook, `${safeId}_${safeName}_텍스트데이터.xlsx`);
  };
  const downloadPdf = async () => {
    if (!window.html2pdf) {
      setSaveError("PDF 라이브러리를 불러오지 못했습니다. 잠시 후 다시 시도해 주세요.");
      setSaveState("error");
      return;
    }
    const report = document.getElementById("lab");
    if (!report) return;
    const now = new Date();
    const date = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}`;
    setPdfMode(true);
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    report.classList.add("pdf-exporting");
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    try {
      await window.html2pdf().set({
        margin: [10, 10, 10, 10],
        filename: `텍스트분석_결과보고서_${date}.pdf`,
        image: { type: "jpeg", quality: 0.96 },
        html2canvas: { scale: 2, useCORS: true, backgroundColor: "#ffffff", scrollX: 0, scrollY: 0 },
        jsPDF: { unit: "mm", format: "a4", orientation: "portrait" },
        pagebreak: { mode: ["css", "legacy"] },
      }).from(report).save();
    } finally {
      report.classList.remove("pdf-exporting");
      setPdfMode(false);
    }
  };
  const cards = ["댓글 입력", "형태소 분석", "불용어 제거", "TF", "IDF", "TF-IDF", "TF-IDF 순위", "벡터 임베딩", "유사도 히트맵", "댓글 군집화"];

  return (
    <main className={motionEnabled ? "motion-on" : "motion-off"}>
      <section className="hero">
        <div className="bread-particles" aria-hidden="true" />
        <nav><div className="brand"><span className="brand-dot">⌁</span> 텍스트 데이터 분석</div><div className="nav-tools"><button className="motion-toggle" onClick={() => setMotionEnabled((enabled) => !enabled)} aria-pressed={motionEnabled}>{motionEnabled ? "✦ 애니메이션 켜짐" : "애니메이션 꺼짐"}</button><div className="nav-note">AI를 읽는 가장 쉬운 방법</div></div></nav>
        <div className="hero-grid">
          <div><p className="eyebrow">AI MATH · TEXT DATA ANALYSIS</p><h1><em>TF-IDF</em>와<br /><em>코사인 유사도</em>로<br />텍스트를 분석해요.</h1><p className="intro">인공지능 수학의 핵심 개념을 활용해 댓글을 단어와 숫자로 바꾸고, 문장 사이의 의미를 비교하는 과정을 직접 체험해요.</p><button className="primary" onClick={() => document.getElementById("lab")?.scrollIntoView({ behavior: "smooth" })}>실험 시작하기 <span>↓</span></button></div>
          <div className="hero-art" aria-hidden="true"><div className="orbit orbit-a"/><div className="orbit orbit-b"/><div className="core"/><span className="float f1">형태소</span><span className="float f2">TF-IDF</span><span className="float f3">유사도</span><span className="sparkle s1">✦</span><span className="sparkle s2">✦</span><span className="sparkle s3">✦</span></div>
        </div>
      </section>

      <section className="lab" id="lab">
        <div className="pdf-header"><h1>인공지능 수학 텍스트 분석 활동 보고서</h1><div><span>작성일시: {savedAt ? new Date(savedAt).toLocaleString("ko-KR") : new Date().toLocaleString("ko-KR")}</span><span>학생 이름: {studentName || "________________"}</span><span>학번: {studentId || "________________"}</span></div></div>
        <div className="section-heading"><div><p className="eyebrow">LEARNING LAB</p><h2>한 단계씩, AI의 생각을 따라가기</h2></div><div className="progress">STEP <b>{step}</b> / 10</div></div>
        <div className="stepper">{cards.map((label, index) => <button key={label} onClick={() => setStep(index + 1)} className={step === index + 1 ? "active" : step > index + 1 ? "done" : ""}><span>{String(index + 1).padStart(2, "0")}</span>{label}</button>)}</div>

        <div className="workbench">
          <aside className="pdf-hide"><p className="side-label">오늘의 미션</p><h3>우리 반의 AI 수업<br/>후기를 분석해 볼까요?</h3><p>댓글을 바꾸면 모든 결과가 즉시 새로 계산돼요.</p><button className="primary" onClick={saveAnalysis} disabled={saveState === "saving"}>{saveState === "saving" ? "제출 중..." : "☁ 분석 결과 제출"}</button><button className="primary pdf-button" onClick={downloadPdf}>▣ PDF 보고서 다운로드</button><button className="primary download-button" onClick={downloadExcel}>↧ 엑셀로 다운로드</button>{saveState === "saved" && <p className="save-message success">수학 선생님에게 분석 결과를 제출했어요.</p>}{saveState === "error" && <p className="save-message error">{saveError || "제출에 실패했어요."}</p>}<div className="tip">✦ <span>Okt처럼 문장을 단어 단위로 나누는 원리를 간단히 체험합니다.</span></div></aside>
          <div className={`stage ${motionEnabled ? "motion-on" : "motion-off"}`}>
            {step === 1 && <div className="panel"><PanelTitle n="01" title="학생 정보를 입력해 주세요" text="이름과 학번은 분석 결과와 함께 Google Sheets에 저장됩니다." /><label className="comment"><span>학생 이름</span><input value={studentName} onChange={(e) => setStudentName(e.target.value)} placeholder="예: 김민영" /></label><label className="comment"><span>학번</span><input value={studentId} onChange={(e) => setStudentId(e.target.value)} placeholder="예: 20101" inputMode="numeric" /></label><PanelTitle n="02" title="댓글을 입력해 주세요" text="서로 다른 문장들을 비교하면 AI가 공통점과 차이를 더 잘 찾아낼 수 있어요." />{comments.map((comment, i) => <label className="comment" key={i}><span>댓글 {i + 1}</span><textarea value={comment} onChange={(e) => update(i, e.target.value)} /></label>)}<button className="add" onClick={() => setComments([...comments, "새로운 댓글을 입력해 보세요."])}>+ 댓글 추가</button></div>}
            {step === 2 && <div className="panel"><PanelTitle n="02" title="형태소 분석 결과" text="Okt 형태소 분석기는 문장을 의미 있는 단어 조각으로 나눠요." />{data.raw.map((words, i) => <ResultRow key={i} index={i} words={words} color="mint" />)}</div>}
            {step === 3 && <div className="panel"><PanelTitle n="03" title="불용어를 걸러냈어요" text="조사처럼 자주 나오지만 의미를 구별하기 어려운 단어를 제거하고, 친절하셔서·친절했습니다·친절한·친절함처럼 같은 뜻의 활용형은 친절로 묶어요." />{data.raw.map((words, i) => <div className="result-row" key={i}><b>댓글 {i + 1}</b><div>{words.map((word, x) => <span className={stopwords.has(word) ? "token removed" : "token mint"} key={x}>{stopwords.has(word) ? word : normalizeWord(word)}</span>)}</div></div>)}</div>}
            {step === 4 && <div className="panel"><div className="weight-direction tf-direction">TF <span>↑</span><small>반복 횟수가 많을수록 TF-IDF에 더 큰 영향을 줘요</small></div><PanelTitle n="04" title="TF: 단어 빈도를 세어요" text="TF(Term Frequency)는 한 댓글 안에서 단어가 몇 번 반복되었는지 보여줘요. 같은 댓글에서 2번 이상 반복된 단어는 주황색으로 강조했어요." />{data.cleaned.map((doc, i) => <div className="bar-row" key={i}><b>댓글 {i + 1}</b><div>{[...new Set(doc)].map((word) => { const count = doc.filter((w) => w === word).length; const isRepeated = count >= 2; return <span className={`bar ${isRepeated ? "repeated" : "single"} ${selectedCommon === word ? "selected-word" : ""}`} key={word} onClick={() => data.df[word] >= 2 && setSelectedCommon(word)} title={`${word}: 이 댓글에서 ${count}번 반복${isRepeated ? "된 단어" : "된 단어"}`} style={{ width: `${Math.max(64, count * 72)}px` }}>{word} <small>×{count}</small></span>; })}</div></div>)}</div>}
            {step === 5 && <div className="panel"><div className="weight-direction idf-direction">IDF <span>↓</span><small>많은 댓글에 등장할수록 IDF 가중치는 낮아져요</small></div><PanelTitle n="05" title="IDF: 희소성을 계산해요" text="IDF가 작은 순서로 최대 12개 단어를 정리했어요. 여러 댓글에 등장하는 단어일수록 IDF가 작아져요. IDF = log((문서 수 + 1) / (포함 문서 수 + 1)) + 1" /><table><thead><tr><th>단어</th><th>포함 댓글 수 (DF)</th><th>IDF</th></tr></thead><tbody>{idfRows.map((word) => <tr key={word}><td>{word}</td><td>{data.df[word]}</td><td>{data.idf[word].toFixed(2)}</td></tr>)}</tbody></table></div>}
            {step === 6 ? <div className="panel"><div className="weight-direction combined-direction"><span>TF ↑</span><b>×</b><span>IDF ↓</span><strong>→ TF-IDF</strong><small>반복 빈도와 희소성 가중치가 곱해져 의미의 좌표가 돼요</small></div><PanelTitle n="06" title="TF-IDF: 의미의 좌표 만들기" text="자주 나오면서도 다른 댓글에는 드문 단어에 더 높은 점수를 줘요. 주황색은 여러 댓글에 공통으로 등장한 단어, 짙은 코코아색은 한 댓글에만 등장해 그 댓글의 특징을 보여주는 단어예요. 공통 단어를 누르면 댓글 위치가 함께 빛나요." /><div className="filter-row" role="group" aria-label="TF-IDF 차트 보기 옵션"><button className={tfidfFilter === "all" ? "selected" : ""} onClick={() => setTfidfFilter("all")}>전체 단어 보기</button><button className={tfidfFilter === "common" ? "selected" : ""} onClick={() => setTfidfFilter("common")}>공통 단어만 보기</button><button className={tfidfFilter === "top5" ? "selected" : ""} onClick={() => setTfidfFilter("top5")}>상위 5개</button><button className={tfidfFilter === "top8" ? "selected" : ""} onClick={() => setTfidfFilter("top8")}>상위 8개</button></div><div className="score-legend"><span><i className="common-swatch" /> 주황색: 여러 댓글에 공통으로 등장</span><span><i className="single-swatch" /> 코코아색: 한 댓글에만 등장</span></div><p className="chart-note">공통 단어 수: <b>{data.tfidfRows.filter((item) => item.documentCount >= 2).length}개</b> · 현재 <b>{visibleTfidfRows.length}개</b> 표시 중</p><div className="score-grid">{visibleTfidfRows.map((item) => { const explanation = item.documentCount >= 2 ? `${item.word} (TF-IDF: ${item.score.toFixed(2)}) → 댓글 ${item.documentCount}개에 공통으로 등장해 여러 댓글을 연결하는 단어예요.` : `${item.word} (TF-IDF: ${item.score.toFixed(2)}) → 한 댓글에만 등장해 그 댓글의 특징을 보여주는 단어예요.`; return <div className={`score-item ${item.documentCount >= 2 ? "common" : "single"} ${selectedCommon === item.word ? "selected-word" : ""}`} key={item.word} title={explanation} onClick={() => item.documentCount >= 2 && setSelectedCommon(selectedCommon === item.word ? "" : item.word)} role={item.documentCount >= 2 ? "button" : undefined} tabIndex={item.documentCount >= 2 ? 0 : undefined}><span>{item.word}</span><b>{item.score.toFixed(2)}</b><i style={{ height: `${35 + item.score * 34}px` }}/><em className="score-tooltip">{explanation}</em></div>; })}</div>{selectedCommon && <p className="selected-readout">✦ <b>{selectedCommon}</b>이(가) 포함된 댓글을 연결해 강조하고 있어요.</p>}{visibleTfidfRows.length === 0 && <p className="empty-filter">공통으로 등장한 단어가 아직 없어요. 전체 단어 보기를 선택해 보세요.</p>}<p className="chart-help">💡 TF-IDF 막대의 색을 함께 읽어 보세요. <b>주황색</b>은 여러 댓글에 공통으로 등장하고, <b>코코아색</b>은 한 댓글에만 등장해 그 댓글의 특징을 보여주는 단어예요.</p></div> : null}
            {(pdfMode || step === 7) && <div className="panel"><PanelTitle n="07" title="TF-IDF 값이 큰 순서로 정리해요" text="각 단어가 댓글을 대표하는 정도를 비교할 수 있도록 TF-IDF가 큰 순서로 최대 12개를 정리했어요." /><table><thead><tr><th>순위</th><th>단어</th><th>가장 큰 TF-IDF</th><th>댓글별 값</th></tr></thead><tbody>{data.tfidfRows.slice(0, 12).map((item, index) => <tr key={item.word}><td>{index + 1}</td><td><b>{item.word}</b></td><td>{item.score.toFixed(2)}</td><td>{item.values.map((value, i) => `댓글 ${i + 1}: ${value.toFixed(2)}`).join(" · ")}</td></tr>)}</tbody></table></div>}
            {(pdfMode || step === 8) && <div className="panel"><PanelTitle n="08" title="표의 값을 벡터로 임베딩해요" text="TF-IDF 점수가 큰 상위 5개 단어만 남겨 5개의 성분을 가진 벡터로 만들어요. 단어를 가로축에 놓고 숫자를 순서대로 배치하면, 댓글 하나가 숫자 벡터가 됩니다." /><p className="embedding-note"><span className="embedding-keywords-line">상위 5개 단어: <b>{data.embeddingRows.map((item) => item.word).join(" · ") || "아직 분석할 단어가 없어요."}</b></span><span className="embedding-note-explanation">데이터가 찾아낸 댓글들의 핵심 키워드, 바로 위 단어들입니다.</span></p><div className="embedding-list">{data.embeddingVectors.map((vector, i) => { const vectorValues = [...vector, 0, 0, 0, 0, 0].slice(0, 5); const vectorName = String.fromCharCode(97 + i); return <div className="embedding-row" key={i}><b>댓글 {i + 1}</b><div className="embedding-scroll"><div className="embedding-words">{data.embeddingRows.map((item) => <span key={item.word}>{item.word}</span>)}</div><div className="embedding-values">{vectorValues.map((value, index) => <span key={`${i}-${index}`}>{value.toFixed(2)}</span>)}</div></div><div className="vector-display"><b>벡터 {vectorName}</b><span className="vector-equals">=</span><code>({vectorValues.map((value) => value.toFixed(2)).join(", ")})</code></div></div>; })}</div></div>}
            {(pdfMode || step === 9) && <div className="panel"><PanelTitle n="09" title="코사인 유사도 히트맵" text="대각선은 자기 자신과의 비교라 흰색으로 표시하고, 나머지는 라벤더(낮은 유사도)에서 골드(높은 유사도)로 읽어요. 가장 비슷한 쌍은 은은한 골드 빛으로 깜박입니다." /><div className="heatmap-layout"><div className="heatmap" style={{ gridTemplateColumns: `repeat(${Math.max(data.matrix.length, 1)}, minmax(0, 1fr))` }}>{data.matrix.flatMap((row, i) => row.map((value, j) => { const diagonal = i === j; const isBest = bestPair ? ((i === bestPair.i && j === bestPair.j) || (i === bestPair.j && j === bestPair.i)) : false; const intensity = Math.max(0, Math.min(1, value)); const low = [174, 190, 226]; const high = [218, 166, 102]; const rgb = low.map((channel, index) => Math.round(channel + (high[index] - channel) * intensity)); return <div key={`${i}-${j}`} className={`${diagonal ? "heatmap-diagonal" : ""} ${isBest ? "heatmap-best" : ""}`} style={{ background: diagonal ? "#fff" : `rgb(${rgb[0]}, ${rgb[1]}, ${rgb[2]})`, color: "#16312c" }}><small>{i + 1} · {j + 1}</small><b>{value.toFixed(2)}</b></div>; }))}</div><div className="cosine-visuals"><figure><img src="/cosine-formula.png" alt="코사인 유사도 공식" /><figcaption>두 벡터의 내적과 크기로 유사도를 계산해요.</figcaption></figure><figure><img src="/cosine-angle.png" alt="두 벡터 사이의 각도 세타 그림" /><figcaption>벡터 사이의 각 θ가 작을수록 두 텍스트가 더 유사해요. 따라서 cos θ가 1에 가까울수록 두 텍스트를 유사하다고 판단해요.</figcaption></figure></div></div><div className="heatmap-legend"><span><i className="blue-swatch" /> 라벤더: 낮은 유사도</span><span><i className="red-swatch" /> 골드: 높은 유사도</span><span><i className="white-swatch" /> 자기 자신과의 비교</span></div><p className="readout">가장 비슷한 문장 쌍: <b>댓글 {(bestPair?.i ?? 0) + 1}</b>와 <b>댓글 {(bestPair?.j ?? 1) + 1}</b></p></div>}
            {(pdfMode || step === 10) && <div className="panel"><PanelTitle n="10" title="댓글을 핵심 단어로 군집화해요" text="왼쪽의 댓글을 읽고, TF-IDF 점수가 높은 핵심 단어를 공유하는 댓글끼리 오른쪽 그룹으로 묶어 보았어요. 같은 색으로 표시하고, 가장 많은 댓글이 포함된 그룹만 반짝여요." /><div className="cluster-board"><div className="comment-cards"><h4>1. 댓글 분석</h4>{comments.map((comment, index) => { const groupIndex = clusterGroups.findIndex((group) => group.documentIndexes.includes(index)); const palette = groupIndex >= 0 ? clusterPalette[groupIndex % clusterPalette.length] : undefined; const paletteStyle = palette ? ({ "--cluster-bg": palette.background, "--cluster-border": palette.border } as any) : undefined; return <div className={`comment-card ${groupIndex >= 0 ? `cluster-comment-${groupIndex}` : "cluster-comment-unassigned"} ${groupIndex === largestClusterIndex ? "largest-cluster" : ""}`} style={paletteStyle} data-cluster-group={groupIndex >= 0 ? String.fromCharCode(65 + groupIndex) : "unassigned"} key={index}><span>댓글 {index + 1}</span><p>{comment || "입력된 댓글이 없습니다."}</p></div>; })}</div><div className="cluster-column"><h4>2. 댓글 분류</h4>{clusterGroups.length ? clusterGroups.map((group, groupIndex) => { const palette = clusterPalette[groupIndex % clusterPalette.length]; const paletteStyle = { "--cluster-bg": palette.background, "--cluster-border": palette.border } as any; return <div className={`cluster-card cluster-card-${groupIndex}`} style={paletteStyle} data-cluster-group={String.fromCharCode(65 + groupIndex)} key={group.label}><div className="cluster-heading"><b>{group.label}</b><span>{group.documentIndexes.length}개 댓글</span></div><p className="cluster-keywords">핵심 단어 {group.keywords.map((word) => <em key={word}>{word}</em>)}</p><div className="cluster-members">{group.documentIndexes.map((docIndex) => <span key={docIndex}>댓글 {docIndex + 1}</span>)}</div></div>; }) : <p className="empty-filter">분류할 단어가 아직 없어요. 댓글을 먼저 입력해 주세요.</p>}</div></div><div className="classification-principle"><div className="principle-flow"><div><b>1. 댓글 분석</b><span>문장에서 핵심 단어를 찾아요</span></div><strong>›››</strong><div><b>2. 댓글 분류</b><span>비슷한 단어를 공유하는 댓글끼리 묶어요</span></div></div><p><b>지금 직접 해 본 원리예요!</b> TF-IDF로 댓글을 대표하는 단어를 찾고, 코사인 유사도로 비슷한 댓글을 비교하면 실제 서비스처럼 댓글을 주제별로 분류할 수 있어요.</p></div></div>}
            <div className="panel-footer pdf-hide"><button disabled={step === 1} onClick={() => setStep(step - 1)}>← 이전</button><button className="next" disabled={step === 10} onClick={() => setStep(step + 1)}>다음 단계 →</button></div>
          </div>
        </div>
      </section>
      <footer><div className="brand"><span className="brand-dot">⌁</span> 텍스트랩</div><p>AI 기반 융합교육 방법 · 텍스트 분석 체험 학습</p></footer>
    </main>
  );
}

function PanelTitle({ n, title, text }: { n: string; title: string; text: string }) { return <div className="panel-title"><span>{n}</span><div><h3>{title}</h3><p>{text}</p></div></div>; }
function ResultRow({ index, words, color }: { index: number; words: string[]; color: string }) { return <div className="result-row"><b>댓글 {index + 1}</b><div>{words.map((word, i) => <span className={`token ${color}`} key={i}>{word}<small>/Noun</small></span>)}</div></div>; }
