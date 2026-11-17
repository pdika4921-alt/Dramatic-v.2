const fs = require('fs');
const p = 'D:/Data/dika baru/webbogor - Salin/public/mitratel.html';
let c = fs.readFileSync(p, 'utf8');
const start = c.indexOf('    // Skala font: pastikan SEMUA teks');
const end = c.indexOf('    } catch (e) { /* thumbnail peta opsional, abaikan */ }');
if (start < 0 || end < 0) { console.log('MARKERS_NOT_FOUND', start, end); process.exit(1); }
const newBlock = `    // Skala font: pastikan SEMUA teks (termasuk nama jalan lengkap) muat di gambar
    const margin = Math.round(Math.max(16, canvas.width / 40) * 1.2);
    const mapSize = Math.min(180, Math.round(canvas.width * 0.18));
    // ukur lebar teks terpanjang
    let base = Math.max(16, Math.round(canvas.width / 40));
    ctx.font = 'bold ' + base + 'px Arial, sans-serif';
    let maxTextWidth = 0;
    for (const ln of overlayLines) {
      const w = ctx.measureText(ln).width;
      if (w > maxTextWidth) maxTextWidth = w;
    }
    // area yang tersedia untuk teks = canvas - margin - mapSize - gap
    const gap = Math.round(margin * 0.8);
    const availTextWidth = canvas.width - margin - mapSize - gap - margin;
    while (base > 10 && maxTextWidth > availTextWidth) {
      base -= 1;
      ctx.font = 'bold ' + base + 'px Arial, sans-serif';
      maxTextWidth = 0;
      for (const ln of overlayLines) {
        const w = ctx.measureText(ln).width;
        if (w > maxTextWidth) maxTextWidth = w;
      }
    }
    const lineHeight = Math.round(base * 1.4);

    // Posisi cluster di kanan bawah: peta tepat di SAMPIING (kiri) teks
    const textRightX = canvas.width - margin;
    const mapX = textRightX - maxTextWidth - gap - mapSize;
    const mapY = canvas.height - mapSize - margin;
    const textBottomY = canvas.height - margin;

    ctx.textAlign = 'right';
    ctx.textBaseline = 'bottom';
    ctx.font = 'bold ' + base + 'px Arial, sans-serif';
    ctx.fillStyle = '#ffffff';
    ctx.shadowColor = 'rgba(0,0,0,0.9)';
    ctx.shadowBlur = Math.round(base * 0.4);
    ctx.shadowOffsetX = 2;
    ctx.shadowOffsetY = 2;

    let y = textBottomY;
    for (let i = overlayLines.length - 1; i >= 0; i--) {
      ctx.fillText(overlayLines[i], textRightX, y);
      y -= lineHeight;
    }

    try {
      const mapUrl = 'https://staticmap.openstreetmap.de/staticmap.php?center=' + encodeURIComponent(lat + ',' + lng) + '&zoom=15&size=180x180&markers=' + encodeURIComponent(lat + ',' + lng + ',red-pushpin');
      const mapImg = new Image();
      mapImg.crossOrigin = 'anonymous';
      mapImg.src = mapUrl;
      await new Promise((res, rej) => { mapImg.onload = res; mapImg.onerror = () => rej(new Error('map fail')); });
      const mw = mapSize;
      const mh = mw;
      ctx.shadowColor = 'transparent';
      ctx.shadowBlur = 0;
      ctx.shadowOffsetX = 0;
      ctx.shadowOffsetY = 0;
      ctx.drawImage(mapImg, mapX, mapY, mw, mh);
      ctx.strokeStyle = '#ffffff';
      ctx.lineWidth = 2;
      ctx.strokeRect(mapX, mapY, mw, mh);
    } catch (e) { /* thumbnail peta opsional, abaikan */ }
`;
c = c.slice(0, start) + newBlock + c.slice(end + '    } catch (e) { /* thumbnail peta opsional, abaikan */ }'.length);
fs.writeFileSync(p, c);
console.log('REPLACED_OK');
`