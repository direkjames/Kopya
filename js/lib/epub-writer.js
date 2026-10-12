// Builds a valid EPUB 3 (with EPUB 2 NCX for older readers).
import { getJSZip } from './loader.js';
import { escapeXml } from './util.js';

const CSS = `body{font-family:serif;line-height:1.5;margin:0 4%}
p{margin:0 0 .8em;text-indent:0}
h1,h2,h3{font-family:sans-serif;line-height:1.25;margin:1.4em 0 .6em}
h1{font-size:1.6em}h2{font-size:1.3em}h3{font-size:1.1em}
.c{text-align:center}.li{margin-left:1.2em;text-indent:-1.2em}
.page{text-align:center;margin:1em 0}.page img{max-width:100%;max-height:95vh}
.cover{text-align:center;margin:0;padding:0}.cover img{max-width:100%;max-height:100vh}`;

const xhtml = (title, body) => `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" lang="en" xml:lang="en">
<head><meta charset="UTF-8"/><title>${escapeXml(title)}</title><link rel="stylesheet" type="text/css" href="style.css"/></head>
<body>
${body}
</body>
</html>`;

// chapters: [{ title, html }], images: [{ name, data(Uint8Array), type }], cover: name|null
export async function writeEpub({ title, author, lang = 'en', chapters, images = [], cover = null }) {
  const JSZip = await getJSZip();
  const zip = new JSZip();
  const id = `urn:uuid:${crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36)}`;
  const modified = new Date().toISOString().replace(/\.\d+Z$/, 'Z');

  zip.file('mimetype', 'application/epub+zip', { compression: 'STORE' });
  zip.file('META-INF/container.xml', `<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles></container>`);
  zip.file('OEBPS/style.css', CSS);
  for (const img of images) zip.file(`OEBPS/images/${img.name}`, img.data);

  const spine = [];
  if (cover) {
    zip.file('OEBPS/cover.xhtml', xhtml('Cover', `<div class="cover"><img src="images/${cover}" alt="Cover"/></div>`));
    spine.push({ id: 'cover', href: 'cover.xhtml', title: 'Cover', linear: true });
  }
  chapters.forEach((ch, i) => {
    const href = `ch${String(i + 1).padStart(3, '0')}.xhtml`;
    zip.file(`OEBPS/${href}`, xhtml(ch.title, ch.html));
    spine.push({ id: `ch${i + 1}`, href, title: ch.title });
  });

  const tocItems = spine.filter((s) => s.id !== 'cover');
  zip.file('OEBPS/nav.xhtml', xhtml('Contents', `<nav epub:type="toc" id="toc"><h1>Contents</h1><ol>${tocItems.map((s) => `<li><a href="${s.href}">${escapeXml(s.title)}</a></li>`).join('')}</ol></nav>`));
  zip.file('OEBPS/toc.ncx', `<?xml version="1.0" encoding="UTF-8"?>
<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1"><head><meta name="dtb:uid" content="${id}"/><meta name="dtb:depth" content="1"/></head><docTitle><text>${escapeXml(title)}</text></docTitle><navMap>${tocItems.map((s, i) => `<navPoint id="np${i + 1}" playOrder="${i + 1}"><navLabel><text>${escapeXml(s.title)}</text></navLabel><content src="${s.href}"/></navPoint>`).join('')}</navMap></ncx>`);

  const imgType = (n) => (/\.png$/i.test(n) ? 'image/png' : 'image/jpeg');
  zip.file('OEBPS/content.opf', `<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="bookid">
<metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
<dc:identifier id="bookid">${id}</dc:identifier>
<dc:title>${escapeXml(title)}</dc:title>
${author ? `<dc:creator>${escapeXml(author)}</dc:creator>` : ''}
<dc:language>${escapeXml(lang)}</dc:language>
<meta property="dcterms:modified">${modified}</meta>
${cover ? '<meta name="cover" content="cover-img"/>' : ''}
</metadata>
<manifest>
<item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>
<item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>
<item id="css" href="style.css" media-type="text/css"/>
${images.map((img, i) => `<item id="${img.name === cover ? 'cover-img' : `img${i}`}" href="images/${img.name}" media-type="${imgType(img.name)}"${img.name === cover ? ' properties="cover-image"' : ''}/>`).join('\n')}
${spine.map((s) => `<item id="${s.id}" href="${s.href}" media-type="application/xhtml+xml"/>`).join('\n')}
</manifest>
<spine toc="ncx">
${spine.map((s) => `<itemref idref="${s.id}"/>`).join('\n')}
</spine>
</package>`);

  return zip.generateAsync({ type: 'blob', mimeType: 'application/epub+zip', compression: 'DEFLATE' });
}
