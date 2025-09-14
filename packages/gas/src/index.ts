/* global */
import { validateAndNormalize } from '@gws/core';

type GasEvent = { parameter?: Record<string, string> } | undefined;

function getJstVersionString(): string {
  try {
    const Utilities = (globalThis as any).Utilities;
    if (Utilities && typeof Utilities.formatDate === 'function') {
      return 'ver.' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    }
  } catch {}
  const d = new Date();
  const yyyy = d.getUTCFullYear();
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(d.getUTCDate()).padStart(2, '0');
  const HH = String(d.getUTCHours()).padStart(2, '0');
  const MM = String(d.getUTCMinutes()).padStart(2, '0');
  const SS = String(d.getUTCSeconds()).padStart(2, '0');
  return `ver.${yyyy}${mm}${dd}_${HH}${MM}${SS}`;
}

function renderTemplate(name: string, vars: Record<string, any> = {}) {
  const HtmlService = (globalThis as any).HtmlService;
  const t = HtmlService.createTemplateFromFile(`views/${name}`);
  Object.assign(t, vars);
  return t
    .evaluate()
    .setTitle('📊 データ管理アプリ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Expose GAS globals via IIFE global
(globalThis as any).doGet = function doGet(e: GasEvent) {
  const page = e?.parameter?.page ?? 'main';
  if (page === 'reception') {
    return renderTemplate('reception_form', { versionString: getJstVersionString() });
  }
  if (page === 'structure_form') {
    return renderTemplate('structure_form', { versionString: getJstVersionString() });
  }
  if (page === 'debug') {
    return renderTemplate('debug', { versionString: getJstVersionString() });
  }
  if (page === 'form_builder') {
    return renderTemplate('form_builder', { versionString: getJstVersionString() });
  }
  return renderTemplate('webapp', {
    versionString: getJstVersionString(),
    showFormBuilder: false,
    showDebugLink: false
  });
};

(globalThis as any).saveReceptionData = function saveReceptionData(payload: unknown) {
  const normalized = validateAndNormalize(payload);
  // TODO: Implement LockService + SpreadsheetApp write. Keep thin adapter.
  return { ok: true, data: normalized };
};

// Compatibility stub for existing UI
(globalThis as any).listReceptionIndex = function listReceptionIndex() {
  return [];
};

