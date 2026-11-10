/** UTF-8 decoder that also works on Hermes builds where TextDecoder may not be exposed. */
export class Utf8StreamDecoder {
  private pending = new Uint8Array(0);
  private firstChunk = true;

  decode(input: Uint8Array, final = false): string {
    const bytes = new Uint8Array(this.pending.length + input.length);
    bytes.set(this.pending, 0);
    bytes.set(input, this.pending.length);
    let end = bytes.length;

    if (!final && end) {
      let lead = end - 1;
      while (lead >= 0 && (bytes[lead] & 0xc0) === 0x80) lead--;
      if (lead >= 0) {
        const first = bytes[lead];
        const expected = first >= 0xf0 && first <= 0xf4 ? 4 : first >= 0xe0 ? 3 : first >= 0xc2 ? 2 : 1;
        if (end - lead < expected) end = lead;
      }
    }

    this.pending = bytes.slice(end);
    const points: number[] = [];
    for (let index = 0; index < end;) {
      const a = bytes[index++];
      let point = 0xfffd;
      if (a < 0x80) point = a;
      else if (a >= 0xc2 && a <= 0xdf && index < end) {
        const b = bytes[index];
        if ((b & 0xc0) === 0x80) { index++; point = ((a & 0x1f) << 6) | (b & 0x3f); }
      } else if (a >= 0xe0 && a <= 0xef && index + 1 < end) {
        const b = bytes[index], c = bytes[index + 1];
        if ((b & 0xc0) === 0x80 && (c & 0xc0) === 0x80 && !(a === 0xe0 && b < 0xa0) && !(a === 0xed && b >= 0xa0)) {
          index += 2; point = ((a & 0x0f) << 12) | ((b & 0x3f) << 6) | (c & 0x3f);
        }
      } else if (a >= 0xf0 && a <= 0xf4 && index + 2 < end) {
        const b = bytes[index], c = bytes[index + 1], d = bytes[index + 2];
        if ((b & 0xc0) === 0x80 && (c & 0xc0) === 0x80 && (d & 0xc0) === 0x80 && !(a === 0xf0 && b < 0x90) && !(a === 0xf4 && b >= 0x90)) {
          index += 3; point = ((a & 7) << 18) | ((b & 0x3f) << 12) | ((c & 0x3f) << 6) | (d & 0x3f);
        }
      }
      points.push(point);
    }
    if (this.firstChunk && points[0] === 0xfeff) points.shift();
    this.firstChunk = false;
    let output = '';
    for (let index = 0; index < points.length; index += 8192) output += String.fromCodePoint(...points.slice(index, index + 8192));
    if (final && this.pending.length) {
      output += '\uFFFD'.repeat(this.pending.length);
      this.pending = new Uint8Array(0);
    }
    return output;
  }
}

