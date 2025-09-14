import { z } from 'zod';

// Domain schema: minimal placeholder for reception payload
export const ReceptionFieldSchema = z.object({
  id: z.string().min(1),
  label: z.string().min(1),
  type: z.enum(['text', 'number', 'date', 'choice']).default('text'),
  required: z.boolean().default(false)
});

export const ReceptionPayloadSchema = z.object({
  receivedAt: z.string().optional(),
  fields: z.array(
    z.object({
      key: z.string().min(1),
      value: z.union([z.string(), z.number(), z.boolean(), z.null()]).nullable()
    })
  )
});

export type ReceptionPayload = z.infer<typeof ReceptionPayloadSchema>;

export const RecordSchema = z.record(z.any());

// 9-level header parse → tree model (simplified bootstrap)
export interface StructureNode {
  key: string;
  title: string;
  children: StructureNode[];
}

// Takes 2D header rows (A1:I*) and builds a nested tree using first non-empty cell per level.
export function parseStructure(headers: string[][]): StructureNode[] {
  const roots: StructureNode[] = [];
  const pathStack: StructureNode[] = [];

  for (const row of headers) {
    // Determine the depth by first non-empty from left
    let depth = -1;
    for (let i = 0; i < Math.min(9, row.length); i++) {
      if (row[i] && row[i].toString().trim() !== '') {
        depth = i;
        break;
      }
    }
    if (depth < 0) continue; // skip empty line

    const title = row[depth]!.toString().trim();
    const key = toKey(title);

    // Adjust stack to current depth
    while (pathStack.length > depth) pathStack.pop();

    const node: StructureNode = { key, title, children: [] };
    if (pathStack.length === 0) {
      roots.push(node);
    } else {
      const parent = pathStack[pathStack.length - 1];
      parent.children.push(node);
    }
    pathStack.push(node);
  }

  return roots;
}

function toKey(s: string): string {
  return s
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_|_$/g, '')
    .slice(0, 64);
}

// Validate and normalize reception payload
export function validateAndNormalize(payload: unknown): ReceptionPayload {
  const p = ReceptionPayloadSchema.parse(payload);
  return {
    receivedAt: p.receivedAt ?? new Date().toISOString(),
    fields: p.fields.map((f) => ({ key: f.key.trim(), value: f.value }))
  };
}

