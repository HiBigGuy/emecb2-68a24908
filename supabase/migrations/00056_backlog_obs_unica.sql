-- ============================================================
-- Migration: Observação única por O.S. no Backlog
-- ============================================================

-- Tabela: backlog_obs_unica (observação editável direto na coluna da tabela)
CREATE TABLE IF NOT EXISTS backlog_obs_unica (
  om TEXT PRIMARY KEY,
  obs TEXT,
  atualizado_em TIMESTAMPTZ DEFAULT now()
);
