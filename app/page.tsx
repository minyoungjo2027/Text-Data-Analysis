"use client";

import { useEffect, useMemo, useState } from "react";
import { saveAnalysisToGoogleSheets } from "../lib/google-sheets";

declare global {
  interface Window {
    XLSX?: any;
  }
}

const stopwords = new Set(["은", "는", "이", "가", "을", "를", "에", "의", "와", "과", "도", "그리고", "하지만", "정말", "너무", "있다", "하다", "좋다"]);
const initial = [
  "AI 수업에서 직접 단어를 분석해 보니 신기하고 재미있어요!",
  "텍스트 데이터를 숫자로 바꾸는 과정이 궁금했어요.",
  "친구들과 비슷한 생각을 했는지 비교해 보고 싶어요.",
];
const STORAGE_KEY = "textlab-analysis-session";
type TfidfFilter = "all" | "common" | "top5" | "top8";

function tokens(text: string) {
  return text
    .replace(/[^가-힣a-zA-Z0-9\s]/g, " ")
    .split(/\s+/)
    .filter((word) => word.length > 1)
    .map((word) => word.toLowerCase());
}

export default function Home() {
  const [comments, setComments] = useState(initial);
  const [studentName, setStudentName] = useState("");
  const [studentId, setStudentId] = useState("");
  const [step, setStep] = useState(1);
  const [saveState, setSaveState] = useState<"idle" | "saving" | "saved" | "error">("idle");
  const [saveError, setSaveError] = useState("");
  const [savedAt, setSavedAt] = useState("");
  const [tfidfFilter, setTfidfFilter] = useState<TfidfFilter>("all");

  useEffect(() => {
    const stored = window.localStorage.getItem(STORAGE_KEY);
    if (!stored) return;
    try {
      const session = JSON.parse(stored) as { studentName?: string; studentId?: string; comments?: string[]; savedAt?: string };
      if (session.studentName) setStudentName(session.studentName);
      if (session.studentId) setStudentId(session.studentId);
      if (session.comments?.length) setComments(session.comments);
      if (session.savedAt) setSavedAt(session.savedAt);
    } catch {
      window.localStorage.removeItem(STORAGE_KEY);
    }
  }, []);
  const data = useMemo(() => {
    const raw = comments.map(tokens);
    const cleaned = raw.map((words) => words.filter((word) => !stopwords.has(word)));
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
    return { raw, cleaned, vocabulary, df, idf, vectors, tfidfRows, matrix: vectors.map((a) => vectors.map((b) => cosine(a, b))) };
  }, [comments]);
  const visibleTfidfRows = data.tfidfRows.filter((item, index) => {
    if (tfidfFilter === "common") return item.documentCount >= 2;
    if (tfidfFilter === "top5") return index < 5;
    if (tfidfFilter === "top8") return index < 8;
    return true;
  });

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
    setSavedAt(timestamp);
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify({
      studentName: studentName.trim(),
      studentId: studentId.trim(),
      comments,
      savedAt: timestamp,
    }));
    const result = await saveAnalysisToGoogleSheets({
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
    const headers = ["학번", "이름", "저장 시간", "댓글1", "댓글2", "댓글3", "댓글4", "댓글5", "댓글6", "댓글7"];
    const row = [studentId.trim(), studentName.trim(), savedAt || new Date().toISOString(), ...comments.slice(0, 7)];
    while (row.length < headers.length) row.push("");
    const worksheet = window.XLSX.utils.aoa_to_sheet([headers, row]);
    const workbook = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(workbook, worksheet, "텍스트 데이터");
    const safeId = studentId.trim().replace(/[\\/:*?"<>|]/g, "_");
    const safeName = studentName.trim().replace(/[\\/:*?"<>|]/g, "_");
    window.XLSX.writeFile(workbook, `${safeId}_${safeName}_텍스트데이터.xlsx`);
  };
  const cards = ["댓글 입력", "형태소 분석", "불용어 제거", "TF", "IDF", "TF-IDF", "TF-IDF 순위", "벡터 임베딩", "유사도 히트맵"];

  return (
    <main>
      <section className="hero">
        <nav><div className="brand"><span className="brand-dot">⌁</span> 텍스트랩</div><div className="nav-note">AI를 읽는 가장 쉬운 방법</div></nav>
        <div className="hero-grid">
          <div><p className="eyebrow">TEXT ANALYSIS · PLAYGROUND</p><h1>문장이 <em>데이터</em>가 되는<br />순간을 만나보세요.</h1><p className="intro">댓글 한 줄에서 시작해, AI가 글을 읽고 숫자로 바꾸고 서로의 의미를 비교하는 과정을 직접 따라가요.</p><button className="primary" onClick={() => document.getElementById("lab")?.scrollIntoView({ behavior: "smooth" })}>실험 시작하기 <span>↓</span></button></div>
          <div className="hero-art" aria-hidden="true"><div className="orbit orbit-a"/><div className="orbit orbit-b"/><div className="core">AI<br/><small>TEXT</small></div><span className="float f1">형태소</span><span className="float f2">TF-IDF</span><span className="float f3">유사도</span></div>
        </div>
      </section>

      <section className="lab" id="lab">
        <div className="section-heading"><div><p className="eyebrow">LEARNING LAB</p><h2>한 단계씩, AI의 생각을 따라가기</h2></div><div className="progress">STEP <b>{step}</b> / 9</div></div>
        <div className="stepper">{cards.map((label, index) => <button key={label} onClick={() => setStep(index + 1)} className={step === index + 1 ? "active" : step > index + 1 ? "done" : ""}><span>{String(index + 1).padStart(2, "0")}</span>{label}</button>)}</div>

        <div className="workbench">
          <aside><p className="side-label">오늘의 미션</p><h3>우리 반의 AI 수업<br/>후기를 분석해 볼까요?</h3><p>댓글을 바꾸면 모든 결과가 즉시 새로 계산돼요.</p><button className="primary" onClick={saveAnalysis} disabled={saveState === "saving"}>{saveState === "saving" ? "저장 중..." : "☁ 분석 결과 저장"}</button><button className="primary download-button" onClick={downloadExcel}>↧ 엑셀로 다운로드</button>{saveState === "saved" && <p className="save-message success">Google Sheets와 브라우저에 저장했어요.</p>}{saveState === "error" && <p className="save-message error">{saveError || "저장에 실패했어요."}</p>}<div className="tip">✦ <span>Okt처럼 문장을 단어 단위로 나누는 원리를 간단히 체험합니다.</span></div></aside>
          <div className="stage">
            {step === 1 && <div className="panel"><PanelTitle n="01" title="학생 정보를 입력해 주세요" text="이름과 학번은 분석 결과와 함께 Google Sheets에 저장됩니다." /><label className="comment"><span>학생 이름</span><input value={studentName} onChange={(e) => setStudentName(e.target.value)} placeholder="예: 김민영" /></label><label className="comment"><span>학번</span><input value={studentId} onChange={(e) => setStudentId(e.target.value)} placeholder="예: 20101" inputMode="numeric" /></label><PanelTitle n="02" title="댓글을 입력해 주세요" text="서로 다른 문장 3개를 비교하면 AI가 공통점과 차이를 더 잘 찾아낼 수 있어요." />{comments.map((comment, i) => <label className="comment" key={i}><span>댓글 {i + 1}</span><textarea value={comment} onChange={(e) => update(i, e.target.value)} /></label>)}<button className="add" onClick={() => setComments([...comments, "새로운 댓글을 입력해 보세요."])}>+ 댓글 추가</button></div>}
            {step === 2 && <div className="panel"><PanelTitle n="02" title="형태소 분석 결과" text="Okt 형태소 분석기는 문장을 의미 있는 단어 조각으로 나눠요." />{data.raw.map((words, i) => <ResultRow key={i} index={i} words={words} color="mint" />)}</div>}
            {step === 3 && <div className="panel"><PanelTitle n="03" title="불용어를 걸러냈어요" text="조사처럼 자주 나오지만 의미를 구별하기 어려운 단어를 제거해요." />{data.raw.map((words, i) => <div className="result-row" key={i}><b>댓글 {i + 1}</b><div>{words.map((word, x) => <span className={stopwords.has(word) ? "token removed" : "token mint"} key={x}>{word}</span>)}</div></div>)}</div>}
            {step === 4 && <div className="panel"><PanelTitle n="04" title="TF: 단어 빈도를 세어요" text="TF(Term Frequency)는 한 댓글 안에서 단어가 몇 번 나왔는지 보여줘요. 진한 초록색은 2개 이상 댓글에 공통으로 등장한 단어예요." />{data.cleaned.map((doc, i) => <div className="bar-row" key={i}><b>댓글 {i + 1}</b><div>{[...new Set(doc)].map((word) => { const count = doc.filter((w) => w === word).length; const isCommon = data.df[word] >= 2; return <span className={`bar ${isCommon ? "common" : "single"}`} key={word} title={`${word}: 댓글 ${data.df[word]}개에 등장${isCommon ? "한 공통 단어" : "한 단일 단어"}`} style={{ width: `${Math.max(64, count * 72)}px` }}>{word} <small>×{count}</small></span>; })}</div></div>)}</div>}
            {step === 5 && <div className="panel"><PanelTitle n="05" title="IDF: 희소성을 계산해요" text="여러 댓글에 자주 등장할수록 특별함은 낮아져요. IDF = log((문서 수 + 1) / (포함 문서 수 + 1)) + 1" /><table><thead><tr><th>단어</th><th>포함 댓글 수 (DF)</th><th>IDF</th></tr></thead><tbody>{data.vocabulary.map((word) => <tr key={word}><td>{word}</td><td>{data.df[word]}</td><td>{data.idf[word].toFixed(2)}</td></tr>)}</tbody></table></div>}
            {step === 6 && <div className="panel"><PanelTitle n="06" title="TF-IDF: 의미의 좌표 만들기" text="자주 나오면서도 다른 댓글에는 드문 단어에 더 높은 점수를 줘요. 공통 단어는 진한 색, 한 댓글에만 나온 단어는 연한 색으로 표시했어요." /><div className="filter-row" role="group" aria-label="TF-IDF 차트 보기 옵션"><button className={tfidfFilter === "all" ? "selected" : ""} onClick={() => setTfidfFilter("all")}>전체 단어 보기</button><button className={tfidfFilter === "common" ? "selected" : ""} onClick={() => setTfidfFilter("common")}>공통 단어만 보기</button><button className={tfidfFilter === "top5" ? "selected" : ""} onClick={() => setTfidfFilter("top5")}>상위 5개</button><button className={tfidfFilter === "top8" ? "selected" : ""} onClick={() => setTfidfFilter("top8")}>상위 8개</button></div><p className="chart-note">공통 단어 수: <b>{data.tfidfRows.filter((item) => item.documentCount >= 2).length}개</b> · 현재 <b>{visibleTfidfRows.length}개</b> 표시 중</p><div className="score-grid">{visibleTfidfRows.map((item) => { const explanation = `${item.word} (TF-IDF: ${item.score.toFixed(2)}) → 댓글 ${item.documentCount}개에 공통으로 등장${item.documentCount >= 2 ? "하여 여러 문장을 연결하는 단어예요." : "하여 이 댓글만의 특징을 보여주는 단어예요."}`; return <div className={`score-item ${item.documentCount >= 2 ? "common" : "single"}`} key={item.word} title={explanation}><span>{item.word}</span><b>{item.score.toFixed(2)}</b><i style={{ height: `${35 + item.score * 34}px` }}/><em className="score-tooltip">{explanation}</em></div>; })}</div>{visibleTfidfRows.length === 0 && <p className="empty-filter">공통으로 등장한 단어가 아직 없어요. 전체 단어 보기를 선택해 보세요.</p>}<p className="chart-help">💡 TF가 같아도 여러 댓글에 걸쳐 등장하는 단어와 한 댓글에서만 중요한 단어는 역할이 달라요. 색상과 필터로 두 특징을 비교해 보세요.</p></div>}
            {step === 7 && <div className="panel"><PanelTitle n="07" title="TF-IDF 값이 큰 순서로 정리해요" text="각 단어가 댓글을 대표하는 정도를 비교할 수 있도록 가장 큰 TF-IDF 값부터 나열했어요." /><table><thead><tr><th>순위</th><th>단어</th><th>가장 큰 TF-IDF</th><th>댓글별 값</th></tr></thead><tbody>{data.tfidfRows.map((item, index) => <tr key={item.word}><td>{index + 1}</td><td><b>{item.word}</b></td><td>{item.score.toFixed(2)}</td><td>{item.values.map((value, i) => `댓글 ${i + 1}: ${value.toFixed(2)}`).join(" · ")}</td></tr>)}</tbody></table></div>}
            {step === 8 && <div className="panel"><PanelTitle n="08" title="표의 값을 벡터로 임베딩해요" text="단어를 가로축에 놓고 숫자를 순서대로 배치하면, 댓글 하나가 숫자 벡터가 됩니다." /><div className="embedding-list">{data.vectors.map((vector, i) => <div className="embedding-row" key={i}><b>댓글 {i + 1}</b><div className="embedding-scroll"><div className="embedding-words">{data.vocabulary.map((word) => <span key={word}>{word}</span>)}</div><div className="embedding-values">{vector.map((value, index) => <span key={`${i}-${index}`}>{value.toFixed(2)}</span>)}</div></div></div>)}</div></div>}
            {step === 9 && <div className="panel"><PanelTitle n="09" title="코사인 유사도 히트맵" text="두 댓글의 단어 방향이 비슷할수록 1에 가까워져요." /><div className="heatmap" style={{ gridTemplateColumns: `repeat(${Math.max(data.matrix.length, 1)}, minmax(0, 1fr))` }}>{data.matrix.flatMap((row, i) => row.map((value, j) => <div key={`${i}-${j}`} style={{ background: `rgba(29, 185, 84, ${0.12 + value * 0.88})` }}><small>{i + 1} · {j + 1}</small><b>{value.toFixed(2)}</b></div>))}</div><p className="readout">가장 비슷한 문장 쌍: <b>댓글 {data.matrix.flatMap((row, i) => row.map((v, j) => ({ v, i, j }))).filter((x) => x.i !== x.j).sort((a,b) => b.v-a.v)[0]?.i + 1 || 1}</b>와 <b>댓글 {data.matrix.flatMap((row, i) => row.map((v, j) => ({ v, i, j }))).filter((x) => x.i !== x.j).sort((a,b) => b.v-a.v)[0]?.j + 1 || 2}</b></p></div>}
            <div className="panel-footer"><button disabled={step === 1} onClick={() => setStep(step - 1)}>← 이전</button><button className="next" disabled={step === 9} onClick={() => setStep(step + 1)}>다음 단계 →</button></div>
          </div>
        </div>
      </section>
      <footer><div className="brand"><span className="brand-dot">⌁</span> 텍스트랩</div><p>AI 기반 융합교육 방법 · 텍스트 분석 체험 학습</p></footer>
    </main>
  );
}

function PanelTitle({ n, title, text }: { n: string; title: string; text: string }) { return <div className="panel-title"><span>{n}</span><div><h3>{title}</h3><p>{text}</p></div></div>; }
function ResultRow({ index, words, color }: { index: number; words: string[]; color: string }) { return <div className="result-row"><b>댓글 {index + 1}</b><div>{words.map((word, i) => <span className={`token ${color}`} key={i}>{word}<small>/Noun</small></span>)}</div></div>; }
