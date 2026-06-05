# Protocolo de Parceria — IN-NOVA

Formulário web para geração automática do Protocolo de Parceria em **PDF**, com cabeçalho duplo (logo IN-NOVA à esquerda, logo do parceiro à direita).

## Stack
- **React 18** + **Vite 5** + **Tailwind CSS 3**
- **html2pdf.js** — geração de PDF no browser (sem backend)
- **lucide-react** — ícones

## Setup

```bash
npm install
npm run dev    # http://localhost:5173
```

## Deploy

### Vercel
```bash
npm run build
vercel deploy --prod
```
Ou conecte o repositório GitHub ao vercel.com — deteta Vite automaticamente.

### Netlify
```bash
npm run build
# Arraste dist/ para netlify.com/drop
```

## Ficheiros necessários na raiz (`public/`)
- `innova.png` — logo da In-Nova (aparece no cabeçalho esquerdo do PDF)

O logo do parceiro é carregado diretamente pelo utilizador via upload no formulário.

## Funcionalidades
- 6 passos: Parceiro → Parceria → Opções → Finanças → Datas → Gerar
- Upload do logo do parceiro (PNG/JPG/SVG) incorporado no PDF via base64
- Cláusulas opcionais com toggle: Innovation Week, Open Day, RGPD, Formação
- Validação de campos obrigatórios antes de gerar
- PDF A4 com Times New Roman, formatação fiel ao template original
