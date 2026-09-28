# Roadmap de Evolução — Expert Custos v1 → v3

**Data:** 28 de setembro de 2026  
**Objetivo:** Transformar Expert Custos de prototipo isolado → sistema de análise integrado com OmniMulti  
**Contexto:** Julitago Bioenergia LTDA | Phase 5 — Custeio por Absorção | 290 centros de custo

---

## Status Atual (v1 & v2)

### ✅ v1 (expert-custos.jsx)
- Chat baseado em React Hooks
- System prompt embutido (Plano de Contas, 290 CCs, DRE de jan/2026)
- Vault de conhecimento (gaps, fases, métricas)
- API Claude Sonnet (mock inicial)
- Tópicos rápidos (20 botões contextuais)
- Progress bar visual das 15 fases

### ✅ v2 (ExpertCustos-v2.html / Refatorado.jsx)
- Componentes reutilizáveis (MetricsCard, ChatWindow, QuickTopics)
- useChat + useCustos hooks isolados
- Métricas simuladas + sessionStorage cache
- UX: textarea com contador, loading indicators, error handling
- Footer com disclaimer educacional

### ❌ Limitações Atuais
- **API Claude:** hardcoded, sem rate limiting tratado
- **Dados:** simulados (Jan/2026 não é real-time)
- **ERP:** nenhuma integração com OmniMulti
- **Autenticação:** inexistente
- **Persistência:** apenas sessionStorage (perde ao refresh)
- **Exportação:** sem suporte a download de relatórios
- **Mobile:** não otimizado para tablet/smartphone
- **Segurança:** API key potencialmente exposta

---

## Fase 1 → v3: Fundações (Semanas 1-2)

### 1.1 Arquitetura Backend (Node.js + Express)
```
expert-custos-backend/
├── src/
│  ├── config/
│  │  ├── env.ts (DATABASE_URL, CLAUDE_API_KEY, OMNIMULTI_API)
│  │  ├── omnimulti.ts (conexão com Rel_1200, Rel_859, Rel_59)
│  │  └── db.ts (PostgreSQL com Drizzle ORM)
│  ├── services/
│  │  ├── omnimulti.service.ts (fetch dados do ERP)
│  │  ├── claude.service.ts (chamadas à API Claude)
│  │  ├── metrics.service.ts (cálculo de KPIs)
│  │  └── cache.service.ts (Redis ou in-memory TTL)
│  ├── routers/
│  │  ├── chat.router.ts (POST /api/chat, GET /api/chat/history)
│  │  ├── metrics.router.ts (GET /api/metrics/{centroId})
│  │  ├── imports.router.ts (POST /api/imports/rel1200 → parse Excel)
│  │  └── export.router.ts (GET /api/export/dre?formato=pdf|xlsx)
│  ├── middleware/
│  │  ├── auth.ts (JWT ou OAuth Manus)
│  │  ├── rateLimit.ts (Claude: 5 req/min, OmniMulti: 10 req/min)
│  │  └── errorHandler.ts
│  └── index.ts (Express app)
├── drizzle/
│  └── schema.ts (tabelas: users, chats, metrics, imports, decisions)
└── package.json
```

### 1.2 Database Schema (Drizzle)
```sql
-- Armazenar histórico de chats por usuário
CREATE TABLE chats (
  id UUID PRIMARY KEY,
  userId TEXT NOT NULL,
  centroId INT,
  pergunta TEXT NOT NULL,
  resposta TEXT NOT NULL,
  fontes TEXT[] (Rel_1200, Rel_859, Rel_59),
  timestamp TIMESTAMP DEFAULT NOW(),
  modelo TEXT DEFAULT 'claude-sonnet-4'
);

-- Cache de métricas (TTL 5 min)
CREATE TABLE metrics_cache (
  centroId INT PRIMARY KEY,
  custoDireto DECIMAL(12,2),
  custoIndireto DECIMAL(12,2),
  receita DECIMAL(12,2),
  margem DECIMAL(5,2),
  atualizadoEm TIMESTAMP DEFAULT NOW()
);

-- Auditoria: importações de Rel_1200
CREATE TABLE imports (
  id UUID PRIMARY KEY,
  usuarioId TEXT,
  nomeArquivo TEXT,
  linhasProcessadas INT,
  erros TEXT[],
  criado TIMESTAMP DEFAULT NOW()
);

-- Decisões do usuário (custeio, rateios)
CREATE TABLE decisions (
  id UUID PRIMARY KEY,
  usuarioId TEXT,
  centroId INT,
  tipoDecisao TEXT (rateio_criterio, devolucao_custo, exclusao_linha),
  descricao TEXT,
  impactoEmR$ DECIMAL(12,2),
  criado TIMESTAMP DEFAULT NOW()
);
```

### 1.3 Tipos TypeScript (tRPC)
```typescript
// shared/types.ts
export type MetricaInput = {
  centroId: number;
  periodo?: string; // YYYY-MM
};

export type MetricaOutput = {
  centroId: number;
  descricao: string;
  custoDireto: number;
  custoIndireto: number;
  receita: number;
  margemContrib: number;
  kpis: {
    custoPerKm?: number;
    custoPerTon?: number;
    eficiencia?: number;
  };
  fonte: "omnimulti" | "cache" | "simulado";
};

export type ChatInput = {
  pergunta: string;
  centroId?: number;
  contexto?: "fase5" | "orcamento" | "variacao";
};

export type ChatOutput = {
  respostaId: string;
  resposta: string;
  fontes: string[]; // Rel_1200, Rel_859
  sugestoesRelacionadas: string[];
};
```

---

## Fase 2 → v3: Integração OmniMulti (Semanas 3-4)

### 2.1 Parser de Rel_1200 (Visual Software)
```typescript
// src/services/omnimulti.service.ts

export class OmniMultiService {
  /**
   * Exportar Rel_1200 (Livro Diário por Centro de Custo)
   * Entrada: Excel com seções demarcadas por sentinelas (coluna A)
   * Saída: Array normalizado com [account, cc, valor, descricao]
   */
  async parseRel1200(buffer: Buffer): Promise<CostEntry[]> {
    const workbook = XLSX.read(buffer);
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

    const entries: CostEntry[] = [];
    let currentSection = "desconhecido";

    for (const row of rows) {
      // Detectar sentinela (ex: "SEÇÃO FLORESTAIS" em A1)
      if (row[0]?.includes("SEÇÃO") || row[0]?.includes("OPERAÇÃO")) {
        currentSection = row[0];
        continue;
      }

      // Pular linhas vazias e totalizadores
      if (!row[0] || row[0].toString().toLowerCase().includes("total")) continue;

      entries.push({
        account: row[0],      // 3.01.08.03.2
        cc: row[1],           // 210001
        valor: parseFloat(row[2]),
        descricao: row[3],
        data: new Date(row[4]),
        secao: currentSection,
      });
    }

    return entries;
  }

  /**
   * Sincronizar Rel_859 (Dupl./Triplicata)
   * Base para conciliação do Rel_59
   */
  async fetchRel859(): Promise<Duplicidade[]> {
    // TODO: chamar API OmniMulti SQL-read
    // const result = await sqlPool.query(`
    //   SELECT nf, emissao, cliente, produto, qtd, valor
    //   FROM rel_859
    //   ORDER BY emissao DESC LIMIT 1000
    // `);
  }

  /**
   * Conciliação Rel_59 vs Rel_1200
   * Chave: [nf, serie, cnpj, dt_emissao, produto, qtd, valor]
   */
  async reconcileRel59(candidatos: Candidato[]): Promise<Decision[]> {
    // Chave composta rigorosa (7 campos)
    // Confirmação manual + histórico reversível
  }
}
```

### 2.2 Sync Automático (cron job)
```typescript
// src/jobs/omnimulti-sync.ts

import cron from 'node-cron';

// Executar todo dia às 23h (após fechamento do OmniMulti)
cron.schedule('0 23 * * *', async () => {
  console.log('🔄 Iniciando sync OmniMulti → Expert Custos');

  try {
    const rel1200 = await omnimulti.exportRel1200();
    const rel859 = await omnimulti.fetchRel859();
    const rel59 = await omnimulti.fetchRel59();

    // Normalizar e armazenar em DB
    await db.imports.create({
      nomeArquivo: `Rel_1200_${new Date().toISOString()}`,
      linhasProcessadas: rel1200.length,
      erros: [],
    });

    // Recalcular métricas
    await metricsService.recalculateCustos();

    console.log('✅ Sync concluído');
  } catch (error) {
    console.error('❌ Erro no sync:', error);
    // Enviar alerta (email, Slack)
  }
});
```

### 2.3 tRPC Routers
```typescript
// src/trpc/routers.ts

export const appRouter = router({
  metrics: publicProcedure
    .input(z.object({ centroId: z.number(), periodo: z.string().optional() }))
    .query(async ({ input }) => {
      return await metricsService.getMetricas(input.centroId, input.periodo);
    }),

  chat: publicProcedure
    .input(z.object({ pergunta: z.string(), centroId: z.number().optional() }))
    .mutation(async ({ input, ctx }) => {
      const resposta = await claudeService.analyze(input.pergunta, {
        centroId: input.centroId,
        vault: VAULT_CONTEXT, // Injetar dados reais de DB
      });
      
      // Armazenar em DB para histórico
      await db.chats.create({
        userId: ctx.userId,
        pergunta: input.pergunta,
        resposta: resposta.text,
        fontes: resposta.sources,
      });

      return resposta;
    }),

  exportDRE: publicProcedure
    .input(z.object({ formato: z.enum(['pdf', 'xlsx']), periodo: z.string() }))
    .query(async ({ input }) => {
      const dre = await metricsService.buildDRE(input.periodo);
      return await exportService.render(dre, input.formato);
    }),
});
```

---

## Fase 3 → v3: Frontend Evoluído (Semanas 5-6)

### 3.1 Arquitetura Frontend (React 19 + Vite)
```
expert-custos-frontend/
├── src/
│  ├── components/
│  │  ├── Chat/
│  │  │  ├── ChatWindow.tsx (com persistência)
│  │  │  ├── MessageBubble.tsx
│  │  │  └── InputBar.tsx (suporte markdown)
│  │  ├── Metrics/
│  │  │  ├── MetricsGrid.tsx
│  │  │  ├── MetricsCard.tsx
│  │  │  └── MetricsChart.tsx (gráficos Recharts)
│  │  ├── DRE/
│  │  │  ├── DREView.tsx (tabela analítica)
│  │  │  ├── DREComparativo.tsx (real vs orçado)
│  │  │  └── DREExport.tsx (download PDF/XLSX)
│  │  ├── Imports/
│  │  │  ├── FileUpload.tsx (drag-drop Rel_1200)
│  │  │  ├── PreviewTable.tsx
│  │  │  └── ValidationSummary.tsx
│  │  └── Layout/
│  │     ├── Header.tsx (breadcrumb, user, logout)
│  │     ├── Sidebar.tsx (nav com fases)
│  │     └── Footer.tsx (status sync, links)
│  ├── hooks/
│  │  ├── useChat.ts (com TanStack Query)
│  │  ├── useCustos.ts (fetch metrics com cache)
│  │  ├── useMetrics.ts (agregações)
│  │  └── useLocalStorage.ts (persistência)
│  ├── lib/
│  │  ├── trpc-client.ts (config tRPC)
│  │  ├── api.ts (axios com interceptors)
│  │  └── utils.ts (formatadores)
│  ├── pages/
│  │  ├── Dashboard.tsx
│  │  ├── Analise.tsx (chat + métricas lado a lado)
│  │  ├── DRE.tsx (DRE completa + comparativa)
│  │  ├── Imports.tsx (upload Rel_1200)
│  │  └── Configuracoes.tsx (integrações, cache)
│  ├── styles/
│  │  ├── globals.css (Tailwind + custom vars)
│  │  └── animations.css
│  └── App.tsx
├── package.json
└── tsconfig.json
```

### 3.2 Novo Layout (v3)
```
┌─────────────────────────────────────────────────────────────────┐
│ 🎓 Expert em Custos  [v3 — Integrado OmniMulti]  ⚙️ 👤 📋     │ ← Header
├──────────────────────┬───────────────────────────────────────────┤
│                      │                                           │
│  🏠 Dashboard        │  📊 Metrics em Tempo Real                │
│  📊 Análise (Chat)   │  ┌─────────────────────────────┐        │
│  📈 DRE Gerencial    │  │ Centro: [dropdown] Período:│        │
│  📂 Imports          │  │ [jan/2026 ▼]              │        │
│  ⚙️  Config          │  │ ┌─────────────────────────┐ │        │
│                      │  │ │ Custo Total: R$ 842.400 │ │        │
│ 🔗 Vault:           │  │ │ Margem: 23% ↓-0.8%      │ │        │
│ • 4/15 fases OK     │  │ │ Eficiência: 87% ↑+1.2%  │ │        │
│ • 3 gaps abertos    │  │ └─────────────────────────┘ │        │
│ • Sync: 14:32       │  │ [Gráfico: Custo/CC bar]    │        │
│                      │  └─────────────────────────────┘        │
│ Últimas Consultas:   │                                           │
│ 1. Combustível gap   │  💬 Chat com Claude                      │
│ 2. DRE jan/2026      │  ┌─────────────────────────────┐        │
│ 3. Picador THOR cost │  │ Você: Explique absorção   │        │
│                      │  │                             │        │
│                      │  │ Expert: [resposta markdown] │        │
│                      │  │  Fontes: Rel_1200, Rel_859 │        │
│                      │  │                             │        │
│                      │  │ [input area + send button] │        │
│                      │  └─────────────────────────────┘        │
└──────────────────────┴───────────────────────────────────────────┘
  ↑ Sidebar (200px)      ↑ Main Content (flex: 1)
```

### 3.3 ChatWindow com Persistência
```typescript
// src/components/Chat/ChatWindow.tsx

export function ChatWindow() {
  const [messages, setMessages] = useState<Message[]>(() => {
    // Recuperar do localStorage ao montar
    const saved = localStorage.getItem('chat_history');
    return saved ? JSON.parse(saved) : [];
  });

  // Persistir ao change
  useEffect(() => {
    localStorage.setItem('chat_history', JSON.stringify(messages));
  }, [messages]);

  // Usar React Query + tRPC
  const chatMutation = trpc.chat.useMutation({
    onSuccess: (data) => {
      setMessages(prev => [...prev, {
        role: 'assistant',
        content: data.resposta,
        fontes: data.fontes, // Mostrar badges
      }]);
    },
  });

  return (
    <div className="chat-container">
      <div className="messages">
        {messages.map((msg, i) => (
          <MessageBubble key={i} role={msg.role} content={msg.content}>
            {msg.fontes && (
              <div className="fontes">
                {msg.fontes.map(f => (
                  <badge key={f}>{f}</badge>
                ))}
              </div>
            )}
          </MessageBubble>
        ))}
      </div>
      <InputBar onSend={(q) => chatMutation.mutate({ pergunta: q })} />
    </div>
  );
}
```

### 3.4 Gráficos (Recharts)
```typescript
// src/components/Metrics/MetricsChart.tsx

export function CustosPorCC({ data, periodo }: Props) {
  return (
    <BarChart data={data}>
      <CartesianGrid strokeDasharray="3 3" />
      <XAxis dataKey="nome" />
      <YAxis />
      <Tooltip formatter={(value) => `R$ ${value.toLocaleString('pt-BR')}`} />
      <Legend />
      <Bar dataKey="custoDireto" fill="#10b981" name="Custo Direto" />
      <Bar dataKey="custoIndireto" fill="#3b82f6" name="Custo Indireto" />
    </BarChart>
  );
}

// Comparativo Real vs Orçado
export function ComparativoRealVsOrcado({ periodo }: Props) {
  return (
    <ComposedChart data={data}>
      <Bar dataKey="real" fill="#ef4444" name="Real" />
      <Line dataKey="orcado" stroke="#3b82f6" name="Orçado" />
      <Scatter dataKey="variacao" fill="#f59e0b" name="Desvio %" />
    </ComposedChart>
  );
}
```

---

## Fase 4 → v3: Autenticação & Segurança (Semana 7)

### 4.1 Auth (OAuth Manus ou JWT)
```typescript
// src/middleware/auth.ts

export const authMiddleware = async (req, res, next) => {
  const token = req.headers.authorization?.split('Bearer ')[1];
  
  if (!token) {
    return res.status(401).json({ error: 'Sem autenticação' });
  }

  try {
    const decoded = jwt.verify(token, process.env.JWT_SECRET);
    req.user = decoded;
    next();
  } catch (error) {
    res.status(401).json({ error: 'Token inválido' });
  }
};

// Frontend: interceptor axios
axiosInstance.interceptors.request.use((config) => {
  const token = localStorage.getItem('token');
  if (token) {
    config.headers.Authorization = `Bearer ${token}`;
  }
  return config;
});
```

### 4.2 Rate Limiting (por endpoint)
```typescript
// src/middleware/rateLimit.ts

import rateLimit from 'express-rate-limit';

export const claudeLimiter = rateLimit({
  windowMs: 1 * 60 * 1000, // 1 min
  max: 5,                   // 5 req/min
  message: 'Limite Claude atingido. Aguarde 1 minuto.',
});

export const omnimultiLimiter = rateLimit({
  windowMs: 1 * 60 * 1000, // 1 min
  max: 10,                  // 10 req/min
  message: 'Limite OmniMulti atingido.',
});

// Uso:
app.post('/api/chat', claudeLimiter, (req, res) => { ... });
app.get('/api/metrics', omnimultiLimiter, (req, res) => { ... });
```

---

## Fase 5 → v3: Deploy & Documentação (Semana 8)

### 5.1 Docker
```dockerfile
# Dockerfile.backend
FROM node:20-alpine
WORKDIR /app
COPY package*.json ./
RUN pnpm install
COPY src ./src
EXPOSE 3000
CMD ["pnpm", "dev"]

# Dockerfile.frontend
FROM node:20-alpine AS build
WORKDIR /app
COPY package*.json ./
RUN pnpm install
COPY src ./src
RUN pnpm build

FROM nginx:alpine
COPY --from=build /app/dist /usr/share/nginx/html
COPY nginx.conf /etc/nginx/nginx.conf
EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]
```

### 5.2 docker-compose.yml
```yaml
version: '3.9'

services:
  postgres:
    image: postgres:16-alpine
    environment:
      POSTGRES_DB: expert_custos
      POSTGRES_USER: ${DB_USER}
      POSTGRES_PASSWORD: ${DB_PASSWORD}
    volumes:
      - postgres_data:/var/lib/postgresql/data

  backend:
    build: ./backend
    environment:
      DATABASE_URL: postgresql://${DB_USER}:${DB_PASSWORD}@postgres:5432/expert_custos
      CLAUDE_API_KEY: ${CLAUDE_API_KEY}
      OMNIMULTI_API: ${OMNIMULTI_API}
    ports:
      - "3000:3000"
    depends_on:
      - postgres

  frontend:
    build: ./frontend
    ports:
      - "3001:80"
    depends_on:
      - backend

volumes:
  postgres_data:
```

### 5.3 GitHub Actions (CI/CD)
```yaml
# .github/workflows/deploy.yml

name: Deploy Expert Custos v3

on:
  push:
    branches: [main]

jobs:
  test-build-deploy:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      
      - name: Setup Node
        uses: actions/setup-node@v3
        with:
          node-version: 20
      
      - name: Install deps
        run: pnpm install
      
      - name: Tests
        run: pnpm test
      
      - name: Build
        run: pnpm build
      
      - name: Deploy to Manus Space
        env:
          MANUS_TOKEN: ${{ secrets.MANUS_TOKEN }}
        run: |
          pnpm manus:deploy
```

---

## Fase 6 → v3: Evolução Contínua

### Milestone 1 (v3.0 — MVP)
- [x] Backend tRPC + DB
- [x] Integração OmniMulti (parser Rel_1200)
- [x] Frontend com chat + métricas
- [x] Auth JWT
- [x] Rate limiting
- [x] Docker + deploy

### Milestone 2 (v3.1 — Análise Avançada)
- [ ] DRE gerencial completa (real vs orçado)
- [ ] Comparativo multiperiodal
- [ ] Gráficos Recharts
- [ ] Export PDF (Puppeteer)
- [ ] Mobile responsivo

### Milestone 3 (v3.2 — Automação)
- [ ] Sync automático OmniMulti (cron diário)
- [ ] Alertas (Slack/email) quando gaps > 5%
- [ ] Recomendações automáticas via Claude
- [ ] Dimensões mestres (CC, Conta, Produto)

### Milestone 4 (v3.3 — Business Intelligence)
- [ ] Dashboard Power BI integrado
- [ ] Previsão orçamentária (Prophet)
- [ ] Análise de sensibilidade (what-if)
- [ ] Custeio ABC avançado

### Milestone 5 (v3.4 — ERP Nativo)
- [ ] Integração SQL OmniMulti (read-only)
- [ ] Substituir Excel exports por direct SQL
- [ ] Webhooks de notificação de lançamentos

---

## Stack Recomendado (v3)

| Camada | Tecnologia | Rationale |
|--------|-----------|-----------|
| **Frontend** | React 19 + TypeScript | Reatividade, type safety |
| **Estilo** | Tailwind CSS 4 | Utilities, dark mode built-in |
| **Build** | Vite + SWC | Fast HMR, production optimized |
| **Backend** | Node.js + Express | Familiar, ecosystem rich |
| **API** | tRPC 11 | Type-safe, end-to-end |
| **DB** | PostgreSQL + Drizzle | Relational, type-safe ORM |
| **Cache** | Redis + node-cache | Sub-second metrics |
| **Auth** | JWT + bcrypt | Simples, escalável |
| **Excel** | XLSX + Handlebars | Parse + export templates |
| **Charts** | Recharts | React-native, responsive |
| **PDF** | Puppeteer + Handlebars | Browser-based rendering |
| **Deploy** | Docker + Manus Space | Containerized, managed |
| **CI/CD** | GitHub Actions | Native, integrated |
| **Monitoring** | Sentry + Winston | Error tracking + logs |

---

## Próximas Ações (Semana 1)

1. **Criar repositório `expert-custos-v3`** com boilerplate Node + React
2. **Setup inicial:**
   ```bash
   # Backend
   cd backend
   pnpm init
   pnpm add express typescript @types/node tsx
   pnpm add @trpc/server @trpc/client drizzle-orm postgres dotenv
   
   # Frontend
   cd frontend
   pnpm create vite . --template react-ts
   pnpm add @tanstack/react-query recharts axios @mantine/hooks
   ```

3. **Configurar `.env.example`:**
   ```
   # Backend
   DATABASE_URL=postgresql://user:pass@localhost:5432/expert_custos
   CLAUDE_API_KEY=sk-ant-...
   OMNIMULTI_API=https://api.omnimulti.com.br
   OMNIMULTI_USER=henrique.pacheco
   OMNIMULTI_PASS=***
   JWT_SECRET=seu-secret-aleatorio
   
   # Frontend
   VITE_API_URL=http://localhost:3000
   VITE_TINYURL_API=https://tinyurl.com/api
   ```

4. **Primeira feature: Chat tRPC**
   - Endpoint `/api/trpc/chat.analyze`
   - Input: string (pergunta)
   - Output: { resposta, fontes, timestamp }
   - Testar com query via tRPC client

5. **Segunda feature: Integração OmniMulti**
   - Parser de Rel_1200 (Excel → JSON)
   - Validação de estrutura
   - Armazenar em DB
   - Expor via tRPC: `metrics.getCustosPorCC`

---

**Versão do documento:** 1.0  
**Última atualização:** 28 de setembro de 2026  
**Próxima revisão:** Após conclusão da Fase 1 (semana 2)
