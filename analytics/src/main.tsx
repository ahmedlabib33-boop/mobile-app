import React, { useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  AlertTriangle,
  BarChart3,
  CalendarDays,
  CheckCircle2,
  ClipboardList,
  FileText,
  Loader2,
  Send,
  Sparkles,
  Target,
} from "lucide-react";
import "./styles.css";

type FormState = {
  product_brief: string;
  audience: string;
  launch_date: string;
  constraints: string;
  available_assets: string;
  delay_analysis_context: string;
};

type StreamEvent = {
  type: string;
  message?: string;
  delta?: string;
  output?: string;
  name?: string;
  model?: string;
  item_type?: string;
};

const initialForm: FormState = {
  product_brief:
    "Launch an engineering analytics release that converts XER schedules, Excel delay registers, claim PDFs, and meeting notes into a prioritized release plan for delay analysis and management decisions.",
  audience: "Engineering directors, planning engineers, project controls, and contracts team",
  launch_date: "2026-07-15",
  constraints:
    "Need to preserve contractual uncertainty, no invented dates, limited validated cost records, critical path must be checked.",
  available_assets:
    "Primavera XER baseline, monthly update XML, delay event Excel register, claim PDF, owner correspondence notes",
  delay_analysis_context:
    "Potential steel delivery delay, late IFC drawings, and concurrent MEP constraints need TIA screening.",
};

function parseSseChunk(buffer: string, onEvent: (event: StreamEvent) => void) {
  const blocks = buffer.split("\n\n");
  const rest = blocks.pop() ?? "";
  for (const block of blocks) {
    const line = block.split("\n").find((item) => item.startsWith("data: "));
    if (!line) continue;
    try {
      onEvent(JSON.parse(line.slice(6)));
    } catch {
      onEvent({ type: "error", message: "Unable to parse stream event." });
    }
  }
  return rest;
}

function App() {
  const [form, setForm] = useState<FormState>(initialForm);
  const [events, setEvents] = useState<StreamEvent[]>([]);
  const [answer, setAnswer] = useState("");
  const [isRunning, setIsRunning] = useState(false);
  const [error, setError] = useState("");

  const toolEvents = events.filter((event) => event.type === "tool_progress");
  const statusEvents = events.filter((event) => event.type === "status");
  const finalOutput = events.findLast((event) => event.type === "final")?.output;
  const displayText = finalOutput || answer || "Submit the brief to stream the analytics release plan.";

  const readiness = useMemo(() => {
    const hasBrief = form.product_brief.trim().length > 40;
    const hasAudience = form.audience.trim().length > 6;
    const hasDate = form.launch_date.trim().length > 4;
    const hasAssets = form.available_assets.trim().length > 10;
    return [hasBrief, hasAudience, hasDate, hasAssets].filter(Boolean).length;
  }, [form]);

  async function runAgent() {
    setIsRunning(true);
    setError("");
    setEvents([]);
    setAnswer("");
    try {
      const response = await fetch("/api/agent/stream", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(form),
      });
      if (!response.ok || !response.body) {
        throw new Error(`API request failed: ${response.status}`);
      }
      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let buffer = "";
      while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        buffer = parseSseChunk(buffer, (event) => {
          setEvents((current) => [...current, event]);
          if (event.type === "text_delta" && event.delta) {
            setAnswer((current) => current + event.delta);
          }
          if (event.type === "error") {
            setError(event.message || "Agent stream failed.");
          }
        });
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : "Unexpected stream failure.");
    } finally {
      setIsRunning(false);
    }
  }

  const update = (key: keyof FormState, value: string) => setForm((current) => ({ ...current, [key]: value }));

  return (
    <main className="app-shell">
      <header className="topbar">
        <div>
          <div className="brand-row">
            <div className="brand-mark"><BarChart3 size={20} /></div>
            <h1>analytics</h1>
          </div>
          <p>Engineering release planning agent for analytics, TIA, delay analysis, and launch decisions.</p>
        </div>
        <div className="status-strip">
          <span>{readiness}/4 inputs ready</span>
          <span>Agents SDK stream</span>
        </div>
      </header>

      <section className="workspace">
        <aside className="input-panel">
          <div className="panel-title">
            <Target size={18} />
            <span>Launch inputs</span>
          </div>
          <label>
            Product brief
            <textarea value={form.product_brief} onChange={(event) => update("product_brief", event.target.value)} rows={6} />
          </label>
          <div className="two-col">
            <label>
              Audience
              <input value={form.audience} onChange={(event) => update("audience", event.target.value)} />
            </label>
            <label>
              Launch date
              <input type="date" value={form.launch_date} onChange={(event) => update("launch_date", event.target.value)} />
            </label>
          </div>
          <label>
            Constraints
            <textarea value={form.constraints} onChange={(event) => update("constraints", event.target.value)} rows={3} />
          </label>
          <label>
            Available assets
            <textarea value={form.available_assets} onChange={(event) => update("available_assets", event.target.value)} rows={3} />
          </label>
          <label>
            Delay analysis context
            <textarea value={form.delay_analysis_context} onChange={(event) => update("delay_analysis_context", event.target.value)} rows={3} />
          </label>
          <button className="primary-action" onClick={runAgent} disabled={isRunning}>
            {isRunning ? <Loader2 className="spin" size={18} /> : <Send size={18} />}
            {isRunning ? "Streaming plan" : "Build release plan"}
          </button>
          {error && <div className="error-box">{error}</div>}
        </aside>

        <section className="output-panel">
          <div className="output-header">
            <div>
              <div className="panel-title">
                <Sparkles size={18} />
                <span>Agent response</span>
              </div>
              <p>Prioritized plan, risk register, owner checklist, launch copy, and missing-detail questions.</p>
            </div>
            <div className={isRunning ? "run-state active" : "run-state"}>
              {isRunning ? "Running" : "Ready"}
            </div>
          </div>

          <div className="progress-row">
            <div><CalendarDays size={16} /> Launch date locked</div>
            <div><ClipboardList size={16} /> Tool rubric</div>
            <div><FileText size={16} /> Copy draft</div>
            <div><AlertTriangle size={16} /> Risk register</div>
          </div>

          <div className="content-grid">
            <article className="response-card">
              <pre>{displayText}</pre>
            </article>
            <aside className="activity-card">
              <h2>Tool activity</h2>
              {toolEvents.length === 0 ? (
                <p className="muted">Tool calls will appear here during the stream.</p>
              ) : (
                toolEvents.map((event, index) => (
                  <div className="activity-item" key={`${event.name}-${index}`}>
                    <CheckCircle2 size={16} />
                    <div>
                      <strong>{event.name}</strong>
                      <span>{event.message} {event.item_type ? `(${event.item_type})` : ""}</span>
                    </div>
                  </div>
                ))
              )}
              <h2>Stream status</h2>
              {statusEvents.map((event, index) => (
                <div className="status-item" key={`${event.message}-${index}`}>{event.message || event.model}</div>
              ))}
            </aside>
          </div>
        </section>
      </section>
    </main>
  );
}

createRoot(document.getElementById("root")!).render(<App />);
