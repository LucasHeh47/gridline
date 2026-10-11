import { describe, expect, it } from 'vitest';
import {
  compress,
  decompress,
  decryptProjectData,
  encryptProjectData,
  invokeShortcuts,
  joinAttributes,
  readVbaProject,
  splitAttributes,
  writeVbaProject,
} from '../src/core/io/vbaProject';

const hex = (s: string) => Uint8Array.from(s.split(/\s+/).filter(Boolean).map((h) => parseInt(h, 16)));
const text = (b: Uint8Array) => new TextDecoder().decode(b);
const bytes = (s: string) => new TextEncoder().encode(s);

describe('VBA compression', () => {
  it('decompresses the examples from the specification', () => {
    expect(text(decompress(hex('01 19 B0 00 61 62 63 64 65 66 67 68 00 69 6A 6B 6C 6D 6E 6F 70 00 71 72 73 74 75 76 2E')))).toBe('abcdefghijklmnopqrstuv.');
    expect(
      text(decompress(hex('01 2F B0 00 23 61 61 61 62 63 64 65 82 66 00 70 61 67 68 69 6A 01 38 08 61 6B 6C 00 30 6D 6E 6F 70 06 71 02 70 04 10 72 73 74 75 76 10 77 78 79 7A 00 3C'))),
    ).toBe('#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa');
  });

  it('round-trips text of any size', () => {
    const samples = ['', 'a', 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa', 'Sub Main()\r\n  MsgBox "Hi"\r\nEnd Sub\r\n'.repeat(400)];
    let rnd = 1;
    const noise = new Uint8Array(9000).map(() => (rnd = (rnd * 1103515245 + 12345) & 0x7fffffff) & 0xff);
    for (const s of samples) expect(text(decompress(compress(bytes(s))))).toBe(s);
    expect(Array.from(decompress(compress(noise)).subarray(0, 9000))).toEqual(Array.from(noise));
  });

  it('actually compresses repetitive code', () => {
    const src = bytes('Range("A1").Value = 1\r\n'.repeat(200));
    expect(compress(src).length).toBeLessThan(src.length / 5);
  });
});

describe('vbaProject.bin', () => {
  it('encrypts and decrypts project protection data', () => {
    const id = '{917DED54-440B-4FD1-A5C1-74ACF261E600}';
    for (const seed of [0, 7, 0x42, 0xff]) {
      expect(decryptProjectData(encryptProjectData([0xff], id, seed))).toEqual([0xff]);
      expect(decryptProjectData(encryptProjectData([0, 0, 0, 0], id, seed))).toEqual([0, 0, 0, 0]);
    }
  });

  it('writes a project that reads back', () => {
    const modules = [
      { name: 'ThisWorkbook', kind: 'document' as const, source: joinAttributes('ThisWorkbook', 'document', 'Private Sub Workbook_Open()\r\nEnd Sub') },
      { name: 'Sheet1', kind: 'document' as const, source: joinAttributes('Sheet1', 'document', '') },
      { name: 'Module1', kind: 'standard' as const, source: joinAttributes('Module1', 'standard', 'Sub Main()\r\n  Range("A1") = "café"\r\nEnd Sub', undefined, { main: 'M' }) },
      { name: 'Counter', kind: 'class' as const, source: joinAttributes('Counter', 'class', 'Public Total As Long') },
    ];
    const bin = writeVbaProject({ name: 'VBAProject', modules });
    const back = readVbaProject(bin);
    expect(back.name).toBe('VBAProject');
    expect(back.modules.map((m) => [m.name, m.kind])).toEqual([
      ['ThisWorkbook', 'document'],
      ['Sheet1', 'document'],
      ['Module1', 'standard'],
      ['Counter', 'class'],
    ]);
    const m1 = splitAttributes(back.modules[2].source);
    expect(m1.code).toBe('Sub Main()\r\n  Range("A1") = "café"\r\nEnd Sub');
    expect(m1.header).toEqual(['Attribute VB_Name = "Module1"']);
    expect(invokeShortcuts(m1.procAttrs)).toEqual({ main: 'M' });
    expect(splitAttributes(back.modules[0].source).header[1]).toContain('00020819');

    // Rewriting from the original keeps the project and replaces the code
    const again = writeVbaProject({ name: 'VBAProject', modules: [modules[0], modules[1], { ...modules[2], source: joinAttributes('Module1', 'standard', 'Sub Other()\r\nEnd Sub') }] }, bin);
    const back2 = readVbaProject(again);
    expect(back2.modules.map((m) => m.name)).toEqual(['ThisWorkbook', 'Sheet1', 'Module1']);
    expect(splitAttributes(back2.modules[2].source).code).toBe('Sub Other()\r\nEnd Sub');
  });
});
