const sharp = require('sharp');

async function boxes(bytes, limit) {
  const img = sharp(Buffer.from(bytes), { failOn: 'none' }).flatten({ background: '#ffffff' }).greyscale();
  const { data, info } = await img.raw().toBuffer({ resolveWithObject: true });
  const W = info.width;
  const H = info.height;
  const minRun = Math.max(30, Math.round(W * 0.06));
  const runs = [];
  for (let y = 0; y < H; y++) {
    let start = -1;
    for (let x = 0; x <= W; x++) {
      const dark = x < W && data[y * W + x] < limit;
      if (dark && start < 0) start = x;
      if (!dark && start >= 0) {
        if (x - start >= minRun) runs.push({ y: y, x0: start, x1: x - 1 });
        start = -1;
      }
    }
  }
  const groups = [];
  runs.forEach((r) => {
    const g = groups.filter((q) => Math.abs(q.x0 - r.x0) <= 10 && Math.abs(q.x1 - r.x1) <= 10)[0];
    if (g) { g.y0 = Math.min(g.y0, r.y); g.y1 = Math.max(g.y1, r.y); g.rows++; }
    else groups.push({ x0: r.x0, x1: r.x1, y0: r.y, y1: r.y, rows: 1 });
  });
  const big = groups.filter((g) => g.y1 - g.y0 >= 30 && g.x1 - g.x0 >= 40);
  return big
    .filter((g) => !big.some((o) => o !== g && o.x0 <= g.x0 + 2 && o.x1 >= g.x1 - 2 && o.y0 <= g.y0 + 2 && o.y1 >= g.y1 - 2 && (o.x1 - o.x0) * (o.y1 - o.y0) > (g.x1 - g.x0) * (g.y1 - g.y0)))
    .sort((a, b) => a.x0 - b.x0);
}

async function cropFirmas(bytes) {
  let found = await boxes(bytes, 90);
  if (!found.length) found = await boxes(bytes, 150);
  const out = [];
  for (const g of found.slice(0, 4)) {
    const inset = 4;
    const left = g.x0 + inset;
    const top = g.y0 + inset;
    const width = g.x1 - g.x0 - inset * 2;
    const height = g.y1 - g.y0 - inset * 2;
    if (width < 20 || height < 20) continue;
    const r = await sharp(Buffer.from(bytes), { failOn: 'none' })
      .flatten({ background: '#ffffff' })
      .extract({ left: left, top: top, width: width, height: height })
      .resize({ width: 900, height: 900, fit: 'inside', withoutEnlargement: true })
      .jpeg({ quality: 90 })
      .toBuffer({ resolveWithObject: true });
    out.push({ buf: r.data, w: r.info.width, h: r.info.height });
  }
  return out;
}

async function wholeImage(bytes) {
  let img = sharp(Buffer.from(bytes), { failOn: 'none' }).flatten({ background: '#ffffff' });
  try { img = sharp(await img.trim({ threshold: 25 }).toBuffer()); } catch (e) { img = sharp(Buffer.from(bytes), { failOn: 'none' }).flatten({ background: '#ffffff' }); }
  const r = await img.resize({ width: 900, height: 900, fit: 'inside', withoutEnlargement: true }).jpeg({ quality: 90 }).toBuffer({ resolveWithObject: true });
  return { buf: r.data, w: r.info.width, h: r.info.height };
}

async function firmasFromRes(res) {
  const imgs = (res.firmas && res.firmas.length) ? res.firmas : (res.firma ? [res.firma] : []);
  const out = [];
  for (const im of imgs) {
    let crops = [];
    try { crops = await cropFirmas(im.bytes); } catch (e) { crops = []; }
    if (!crops.length) { try { crops = [await wholeImage(im.bytes)]; } catch (e) { crops = []; } }
    crops.forEach((c) => out.push(c));
  }
  return out.slice(0, 6);
}

module.exports = { cropFirmas: cropFirmas, firmasFromRes: firmasFromRes };
