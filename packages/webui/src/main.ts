type SavePayload = { fields: { key: string; value: unknown }[] };

function isGas(): boolean {
  return typeof (globalThis as any).google?.script?.run !== 'undefined';
}

export async function saveReceptionData(payload: SavePayload): Promise<any> {
  if (isGas()) {
    return new Promise((resolve, reject) => {
      (globalThis as any).google.script.run
        .withSuccessHandler(resolve)
        .withFailureHandler(reject)
        .saveReceptionData(payload);
    });
  }
  // Local mock
  return { ok: true, data: payload };
}

