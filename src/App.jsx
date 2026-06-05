import { useState, useRef } from "react";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { saveAs } from "file-saver";

const PRIMARY = "#0065BD";
const SECONDARY = "#0F172A";
const LIGHT_BG = "#F4F7FB";

const toDateStr = (d) => d.toISOString().split("T")[0];
const splitDate = (val) => {
  const [y, m, d] = val.split("-");
  return {
    day: String(parseInt(d, 10)),
    month: String(parseInt(m, 10)),
    year: y,
  };
};

const today = new Date();
const todayStr = toDateStr(today);

const CHECKED = "☑";
const UNCHECKED = "☐";

const LEVELS = [
  { value: "N2", label: "N2 – Participação Básica", desc: "2 pts/hora" },
  { value: "N3", label: "N3 – Participação Ativa", desc: "3 pts/hora" },
  { value: "N4", label: "N4 – Contribuição Autónoma", desc: "4 pts/hora" },
  {
    value: "N5",
    label: "N5 – Responsabilidade Estruturada",
    desc: "5 pts/hora",
  },
];

function fixSplitTags(xmlStr) {
  return xmlStr.replace(/\{#([\s\S]*?)#\}/g, (match) => {
    const inner = match
      .replace(/<[^>]+>/g, "")
      .replace(/\s+/g, " ")
      .trim();
    return inner;
  });
}

// ── Styles (defined outside component so they're stable) ──────────────────

const inputStyle = (hasError) => ({
  width: "100%",
  padding: "10px 12px",
  borderRadius: 8,
  border: `1.5px solid ${hasError ? "#d32f2f" : "#d0d0d8"}`,
  fontSize: 14,
  marginTop: 4,
  outline: "none",
  boxSizing: "border-box",
  background: hasError ? "#fff5f5" : "white",
  color: SECONDARY,
  fontFamily: "inherit",
  transition: "border-color 0.15s",
});

const labelStyle = {
  display: "block",
  marginTop: 18,
  fontSize: 13,
  fontWeight: 600,
  color: "#444",
};

const errorMsgStyle = {
  color: "#d32f2f",
  fontSize: 12,
  marginTop: 3,
  display: "block",
};

const sectionTitleStyle = {
  color: PRIMARY,
  fontWeight: 700,
  marginTop: 32,
  marginBottom: 8,
  fontSize: 15,
  borderBottom: `2px solid ${PRIMARY}22`,
  paddingBottom: 6,
};

// ── Field component (outside App so it's never recreated) ─────────────────

function Field({
  label,
  name,
  type = "text",
  placeholder = "",
  form,
  errors,
  onChange,
}) {
  return (
    <label style={labelStyle}>
      {label}
      <input
        type={type}
        name={name}
        value={form[name]}
        onChange={onChange}
        placeholder={placeholder}
        style={inputStyle(errors[name])}
        className="premium-input"
      />
      {errors[name] && <span style={errorMsgStyle}>Campo obrigatório</span>}
    </label>
  );
}

// ── Main component ─────────────────────────────────────────────────────────

export default function App() {
  const [startDate, setStartDate] = useState(todayStr);
  const [endDate, setEndDate] = useState(todayStr);
  const errorBannerRef = useRef(null);
  const [errors, setErrors] = useState({});

  const [form, setForm] = useState({
    name: "",
    number: "",
    desciption: "",
    city: "",
    entity: "",
    level: "",
    hours: "",
    online: "",
    organizer_name: "",
    organizer_position: "",
    observations: "",
  });

  const handleChange = (e) => {
    const { name, value } = e.target;
    setForm((p) => ({ ...p, [name]: value }));
    if (errors[name]) setErrors((p) => ({ ...p, [name]: false }));
  };

  const required = [
    "name",
    "number",
    "desciption",
    "city",
    "entity",
    "level",
    "hours",
    "online",
    "organizer_name",
    "organizer_position",
  ];

  const validate = () => {
    const newErrors = {};
    required.forEach((k) => {
      if (!form[k] || form[k].trim() === "") newErrors[k] = true;
    });

    if (!startDate) newErrors.startDate = true;
    if (!endDate) newErrors.endDate = true;
    if (startDate && endDate && new Date(endDate) < new Date(startDate)) {
      newErrors.endDate = "Data de fim anterior à data de início";
    }

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const generateDoc = async () => {
    if (!validate()) {
      setTimeout(
        () =>
          errorBannerRef.current?.scrollIntoView({
            behavior: "smooth",
            block: "start",
          }),
        50,
      );
      return;
    }

    const response = await fetch("/template.docx");
    const arrayBuffer = await response.arrayBuffer();

    const zip = new PizZip(arrayBuffer);
    const docXmlKey = "word/document.xml";
    let xmlStr = zip.files[docXmlKey].asText();
    xmlStr = fixSplitTags(xmlStr);
    zip.file(docXmlKey, xmlStr);

    const doc = new Docxtemplater(zip, {
      delimiters: { start: "{#", end: "#}" },
      paragraphLoop: true,
      linebreaks: true,
    });

    const sd = splitDate(startDate);
    const ed = splitDate(endDate);

    const data = {
      name: form.name,
      number: form.number,
      desciption: form.desciption,
      city: form.city,
      entity: form.entity,
      N2: form.level === "N2" ? CHECKED : UNCHECKED,
      N3: form.level === "N3" ? CHECKED : UNCHECKED,
      N4: form.level === "N4" ? CHECKED : UNCHECKED,
      N5: form.level === "N5" ? CHECKED : UNCHECKED,
      st_day: sd.day,
      st_month: sd.month,
      st_year: sd.year,
      end_day: ed.day,
      end_month: ed.month,
      end_year: ed.year,
      hours: form.hours,
      online: form["online"],
      organizer_name: form.organizer_name,
      organizer_position: form.organizer_position,
      observations: form.observations && form.observations.trim() ? form.observations : " ",
    };

    try {
      doc.render(data);
    } catch (err) {
      console.error(err);
      alert(
        "Erro ao preencher o template. Verifique os campos e tente novamente.",
      );
      return;
    }

    const blob = doc.getZip().generate({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });

    saveAs(
      blob,
      `Certificado_${form.name.replace(/\s+/g, "_") || "Participacao"}.docx`,
    );
  };

  const errorCount = Object.keys(errors).length;
  const fieldProps = { form, errors, onChange: handleChange };

  return (
    <div
      style={{
        minHeight: "100vh",
        background: `linear-gradient(135deg, ${LIGHT_BG} 0%, #EBF5FF 100%)`,
        display: "flex",
        justifyContent: "center",
        padding: "40px 20px",
        fontFamily: "'Inter', 'Segoe UI', sans-serif",
      }}
    >
      <div
        style={{
          width: "100%",
          maxWidth: 720,
          background: "white",
          borderRadius: 16,
          padding: "40px 44px",
          boxShadow:
            "0 10px 30px rgba(0, 101, 189, 0.06), 0 1px 3px rgba(0, 0, 0, 0.02)",
          borderTop: `6px solid ${PRIMARY}`,
        }}
      >
        {/* Header */}
        <div style={{ marginBottom: 28 }}>
          <h1
            style={{
              color: SECONDARY,
              margin: 0,
              fontSize: 24,
              fontWeight: 700,
            }}
          >
            Certificado de Participação
          </h1>
          <p style={{ color: "#888", marginTop: 6, fontSize: 14 }}>
            Atividades Complementares A e B — preencha os dados para gerar o
            documento.
          </p>
        </div>

        {/* Error banner */}
        {errorCount > 0 && (
          <div
            ref={errorBannerRef}
            style={{
              background: "#fff5f5",
              border: "1.5px solid #d32f2f",
              borderRadius: 10,
              padding: "12px 16px",
              marginBottom: 24,
              color: "#d32f2f",
              fontSize: 13,
            }}
          >
            <strong>Atenção:</strong> Preencha todos os campos obrigatórios
            antes de gerar o documento.
            {errorCount > 1 && ` (${errorCount} campos em falta)`}
          </div>
        )}

        {/* ── Estudante ── */}
        <div style={sectionTitleStyle}>Dados do Estudante</div>
        <Field
          label="Nome do estudante *"
          name="name"
          placeholder="Nome completo"
          {...fieldProps}
        />
        <Field
          label="Número *"
          name="number"
          placeholder="Ex: 12345"
          {...fieldProps}
        />

        {/* ── Atividade ── */}
        <div style={sectionTitleStyle}>Identificação da Atividade</div>
        <label style={labelStyle}>
          Título e descrição da atividade *
          <textarea
            name="desciption"
            value={form.desciption}
            onChange={handleChange}
            rows={3}
            placeholder="Nome do evento, projeto, curso, ação ou função desempenhada"
            style={{
              ...inputStyle(errors.desciption),
              resize: "vertical",
              lineHeight: 1.5,
            }}
            className="premium-input"
          />
          {errors.desciption && (
            <span style={errorMsgStyle}>Campo obrigatório</span>
          )}
        </label>
        <div
          style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}
        >
          <Field
            label="Cidade *"
            name="city"
            placeholder="Ex: Lisboa"
            {...fieldProps}
          />
          <Field
            label="Entidade organizadora *"
            name="entity"
            placeholder="Ex: NOVA FCT"
            {...fieldProps}
          />
        </div>

        {/* ── Nível ── */}
        <div style={sectionTitleStyle}>Função Desempenhada</div>
        <label style={{ ...labelStyle, marginTop: 4 }}>
          Nível de participação *
        </label>
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "1fr 1fr",
            gap: 10,
            marginTop: 8,
            padding: errors.level ? "10px" : 0,
            border: errors.level
              ? "1.5px solid #d32f2f"
              : "1.5px solid transparent",
            borderRadius: 10,
            background: errors.level ? "#fff5f5" : "transparent",
          }}
        >
          {LEVELS.map(({ value, label, desc }) => {
            const sel = form.level === value;
            return (
              <label
                key={value}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 10,
                  padding: "12px 14px",
                  borderRadius: 8,
                  cursor: "pointer",
                  border: `1.5px solid ${sel ? PRIMARY : "#d0d0d8"}`,
                  background: sel ? PRIMARY + "0f" : "white",
                  transition: "all 0.15s",
                }}
                onMouseOver={(e) => {
                  if (!sel) e.currentTarget.style.borderColor = PRIMARY + "80";
                }}
                onMouseOut={(e) => {
                  if (!sel) e.currentTarget.style.borderColor = "#d0d0d8";
                }}
              >
                <span
                  style={{
                    fontSize: 20,
                    color: sel ? PRIMARY : "#aaa",
                    lineHeight: 1,
                    flexShrink: 0,
                  }}
                >
                  {sel ? CHECKED : UNCHECKED}
                </span>
                <input
                  type="radio"
                  name="level"
                  value={value}
                  checked={sel}
                  onChange={handleChange}
                  style={{ display: "none" }}
                />
                <span>
                  <span
                    style={{
                      fontWeight: 600,
                      fontSize: 13,
                      color: sel ? PRIMARY : SECONDARY,
                    }}
                  >
                    {label}
                  </span>
                  <br />
                  <span style={{ fontSize: 11, color: "#888" }}>{desc}</span>
                </span>
              </label>
            );
          })}
        </div>
        {errors.level && <span style={errorMsgStyle}>Selecione uma opção</span>}

        {/* ── Período ── */}
        <div style={sectionTitleStyle}>Período de Realização</div>
        <div
          style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}
        >
          <label style={labelStyle}>
            Data de início *
            <input
              type="date"
              value={startDate}
              onChange={(e) => {
                setStartDate(e.target.value);
                if (errors.startDate)
                  setErrors((p) => ({ ...p, startDate: false }));
                if (errors.endDate === "Data de fim anterior à data de início")
                  setErrors((p) => ({ ...p, endDate: false }));
              }}
              style={inputStyle(!!errors.startDate)}
              className="premium-input"
            />
            {errors.startDate && (
              <span style={errorMsgStyle}>
                {errors.startDate === true
                  ? "Campo obrigatório"
                  : errors.startDate}
              </span>
            )}
          </label>
          <label style={labelStyle}>
            Data de fim *
            <input
              type="date"
              value={endDate}
              onChange={(e) => {
                setEndDate(e.target.value);
                if (errors.endDate)
                  setErrors((p) => ({ ...p, endDate: false }));
              }}
              style={inputStyle(!!errors.endDate)}
              className="premium-input"
            />
            {errors.endDate && (
              <span style={errorMsgStyle}>
                {errors.endDate === true ? "Campo obrigatório" : errors.endDate}
              </span>
            )}
          </label>
        </div>

        {/* ── Horas ── */}
        <div style={sectionTitleStyle}>Horas e Modalidade</div>
        <div
          style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}
        >
          <Field
            label="Total de horas de contacto *"
            name="hours"
            placeholder="Ex: 8"
            {...fieldProps}
          />
          <label style={labelStyle}>
            Modalidade *
            <select
              name="online"
              value={form["online"]}
              onChange={handleChange}
              style={{ ...inputStyle(errors.online), appearance: "auto" }}
              className="premium-input"
            >
              <option value="">Selecione...</option>
              <option value="(presencial)">(presencial)</option>
              <option value="(online)">(online)</option>
              <option value="(presencial e online)">
                (presencial e online)
              </option>
            </select>
            {errors.online && (
              <span style={errorMsgStyle}>Campo obrigatório</span>
            )}
          </label>
        </div>

        {/* ── Responsável pela Organização ── */}
        <div style={sectionTitleStyle}>Responsável pela Organização</div>
        <div
          style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}
        >
          <Field
            label="Nome do responsável *"
            name="organizer_name"
            placeholder="Ex: Prof. Doutor João Silva"
            {...fieldProps}
          />
          <Field
            label="Cargo / Função *"
            name="organizer_position"
            placeholder="Ex: Coordenador de Curso"
            {...fieldProps}
          />
        </div>

        {/* ── Observações ── */}
        <div style={sectionTitleStyle}>Observações</div>
        <label style={labelStyle}>
          Observações adicionais
          <textarea
            name="observations"
            value={form.observations}
            onChange={handleChange}
            rows={3}
            placeholder="Informações adicionais relevantes..."
            style={{
              ...inputStyle(false),
              resize: "vertical",
              lineHeight: 1.5,
            }}
            className="premium-input"
          />
        </label>

        {/* ── Botão ── */}
        <button
          onClick={generateDoc}
          style={{
            marginTop: 36,
            width: "100%",
            padding: "14px",
            background: PRIMARY,
            color: "white",
            border: "none",
            borderRadius: 10,
            fontSize: 15,
            fontWeight: 700,
            cursor: "pointer",
            letterSpacing: 0.3,
            transition: "all 0.2s ease-in-out",
          }}
          onMouseOver={(e) => {
            e.currentTarget.style.background = "#00529B";
            e.currentTarget.style.boxShadow =
              "0 4px 12px rgba(0, 101, 189, 0.3)";
          }}
          onMouseOut={(e) => {
            e.currentTarget.style.background = PRIMARY;
            e.currentTarget.style.boxShadow = "none";
          }}
          onMouseDown={(e) => {
            e.currentTarget.style.transform = "scale(0.985)";
          }}
          onMouseUp={(e) => {
            e.currentTarget.style.transform = "scale(1)";
          }}
        >
          Gerar Certificado
        </button>
      </div>
    </div>
  );
}
