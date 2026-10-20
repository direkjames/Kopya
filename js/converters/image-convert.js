import { baseName, tick } from '../lib/util.js';
import { FORMATS, readImage, targetSize, renderImage, encodeImage } from '../lib/imaging.js';

const LEVELS = {
  high: { q: 0.92, avif: 80 },
  balanced: { q: 0.8, avif: 60 },
  small: { q: 0.6, avif: 40 },
};

export async function convert(file, { report, options }) {
  const format = FORMATS[options.format] ? options.format : 'jpg';
  const f = FORMATS[format];
  const level = LEVELS[options.quality] || LEVELS.high;
  const maxSide = Number(options.size) || 0;

  report(0.05, 'Opening the image');
  const img = await readImage(file);
  const warnings = [...img.warnings];
  try {
    await tick();
    const size = targetSize(img, maxSide, format);
    warnings.push(...size.warnings);

    report(0.3, 'Drawing the picture');
    // JPG and BMP can't be see-through, so transparent areas become white.
    const out = renderImage(img, size.width, size.height, f.opaque ? '#ffffff' : null);
    if (out.note) warnings.push(out.note);
    await tick();

    report(0.45, `Saving as ${f.label}`);
    const blob = await encodeImage(out.canvas, format, level, report);
    out.canvas.width = out.canvas.height = 0;

    const sameExt = file.name.toLowerCase().endsWith(f.ext) || (format === 'jpg' && /\.jpe?g$/i.test(file.name));
    const name = `${baseName(file.name)}${sameExt ? '-converted' : ''}${f.ext}`;
    return { blob, name, detail: `${out.width} × ${out.height}`, warnings };
  } finally {
    img.close();
  }
}
