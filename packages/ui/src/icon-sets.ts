import type { IconSetName } from '@sheets/core';

const RED = '#C0392B';
const ORANGE = '#E67E22';
const YELLOW = '#E5B100';
const GREEN = '#2E9E44';
const GRAY = '#7F7F7F';

/** Traffic-light color for icon `index` of `n` (lowest = red, highest = green). */
function ramp(index: number, n: number): string {
  if (n <= 3) return [RED, YELLOW, GREEN][index]!;
  if (n === 4) return [RED, ORANGE, YELLOW, GREEN][index]!;
  return [RED, ORANGE, YELLOW, '#8BC34A', GREEN][index]!;
}

function arrow(ctx: CanvasRenderingContext2D, cx: number, cy: number, r: number, angle: number, color: string): void {
  ctx.save();
  ctx.translate(cx, cy);
  ctx.rotate(angle);
  ctx.fillStyle = color;
  ctx.beginPath();
  ctx.moveTo(0, -r);
  ctx.lineTo(r * 0.9, 0);
  ctx.lineTo(r * 0.35, 0);
  ctx.lineTo(r * 0.35, r);
  ctx.lineTo(-r * 0.35, r);
  ctx.lineTo(-r * 0.35, 0);
  ctx.lineTo(-r * 0.9, 0);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}

function star(ctx: CanvasRenderingContext2D, cx: number, cy: number, r: number): void {
  ctx.beginPath();
  for (let i = 0; i < 10; i++) {
    const rad = i % 2 ? r * 0.45 : r;
    const a = -Math.PI / 2 + (i * Math.PI) / 5;
    ctx.lineTo(cx + rad * Math.cos(a), cy + rad * Math.sin(a));
  }
  ctx.closePath();
}

/** Draws one icon of an Excel icon set into a square at (x, y). */
export function drawIcon(ctx: CanvasRenderingContext2D, set: IconSetName, index: number, x: number, y: number, size: number): void {
  const n = Number(set[0]);
  const kind = set.slice(1);
  const cx = x + size / 2, cy = y + size / 2, r = size / 2 - 1;
  const color = ramp(index, n);
  ctx.save();
  ctx.lineWidth = Math.max(1, size / 9);
  switch (kind) {
    case 'Arrows':
    case 'ArrowsGray': {
      // Angles from down (lowest) to up (highest).
      const angles = n === 3 ? [Math.PI, Math.PI / 2, 0] : n === 4 ? [Math.PI, (3 * Math.PI) / 4, Math.PI / 4, 0] : [Math.PI, (3 * Math.PI) / 4, Math.PI / 2, Math.PI / 4, 0];
      arrow(ctx, cx, cy, r, angles[index]!, kind === 'ArrowsGray' ? GRAY : color);
      break;
    }
    case 'Triangles':
      ctx.fillStyle = color;
      ctx.beginPath();
      if (index === 1) ctx.rect(x + 2, cy - size / 10, size - 4, size / 5);
      else if (index === 0) { ctx.moveTo(x + 1, y + size * 0.3); ctx.lineTo(x + size - 1, y + size * 0.3); ctx.lineTo(cx, y + size * 0.8); }
      else { ctx.moveTo(x + 1, y + size * 0.7); ctx.lineTo(x + size - 1, y + size * 0.7); ctx.lineTo(cx, y + size * 0.2); }
      ctx.fill();
      break;
    case 'Flags':
      ctx.fillStyle = color;
      ctx.fillRect(x + size * 0.25, y + 1, size * 0.1, size - 2);
      ctx.beginPath();
      ctx.moveTo(x + size * 0.35, y + 1);
      ctx.lineTo(x + size - 1, y + size * 0.3);
      ctx.lineTo(x + size * 0.35, y + size * 0.6);
      ctx.fill();
      break;
    case 'Signs':
      ctx.fillStyle = color;
      ctx.beginPath();
      if (index === 0) { ctx.moveTo(cx, y + 1); ctx.lineTo(x + size - 1, cy); ctx.lineTo(cx, y + size - 1); ctx.lineTo(x + 1, cy); }
      else if (index === 1) { ctx.moveTo(cx, y + 1); ctx.lineTo(x + size - 1, y + size - 2); ctx.lineTo(x + 1, y + size - 2); }
      else ctx.arc(cx, cy, r, 0, Math.PI * 2);
      ctx.fill();
      break;
    case 'Symbols':
    case 'Symbols2': {
      if (kind === 'Symbols') {
        ctx.fillStyle = color;
        ctx.beginPath();
        ctx.arc(cx, cy, r, 0, Math.PI * 2);
        ctx.fill();
      }
      ctx.strokeStyle = kind === 'Symbols' ? '#fff' : color;
      ctx.fillStyle = ctx.strokeStyle;
      ctx.lineWidth = Math.max(1.5, size / 6);
      ctx.lineCap = 'round';
      const d = r * 0.45;
      ctx.beginPath();
      if (index === 1) {
        // Exclamation mark: a bar and a separate dot.
        ctx.moveTo(cx, cy - d * 1.3);
        ctx.lineTo(cx, cy + d * 0.1);
        ctx.stroke();
        ctx.beginPath();
        ctx.arc(cx, cy + d * 1.05, ctx.lineWidth * 0.55, 0, Math.PI * 2);
        ctx.fill();
        break;
      }
      if (index === 0) { ctx.moveTo(cx - d, cy - d); ctx.lineTo(cx + d, cy + d); ctx.moveTo(cx + d, cy - d); ctx.lineTo(cx - d, cy + d); }
      else { ctx.moveTo(cx - d, cy); ctx.lineTo(cx - d * 0.2, cy + d * 0.8); ctx.lineTo(cx + d, cy - d * 0.8); }
      ctx.stroke();
      break;
    }
    case 'Stars':
      ctx.strokeStyle = YELLOW;
      ctx.fillStyle = YELLOW;
      star(ctx, cx, cy, r);
      ctx.stroke();
      if (index > 0) {
        ctx.save();
        star(ctx, cx, cy, r);
        ctx.clip();
        ctx.fillRect(x, y, index === 1 ? size / 2 : size, size);
        ctx.restore();
      }
      break;
    case 'Rating':
      for (let i = 0; i < n - 1; i++) {
        const h = ((i + 1) / (n - 1)) * (size - 2);
        ctx.fillStyle = i < index ? '#2F6EBA' : '#C9D7EA';
        ctx.fillRect(x + 1 + (i * size) / (n - 1), y + size - 1 - h, size / (n - 1) - 1.5, h);
      }
      break;
    case 'Quarters':
      ctx.strokeStyle = GRAY;
      ctx.fillStyle = GRAY;
      ctx.beginPath();
      ctx.arc(cx, cy, r, 0, Math.PI * 2);
      ctx.stroke();
      if (index > 0) {
        ctx.beginPath();
        ctx.moveTo(cx, cy);
        ctx.arc(cx, cy, r, -Math.PI / 2, -Math.PI / 2 + (Math.PI * 2 * index) / (n - 1));
        ctx.closePath();
        ctx.fill();
      }
      break;
    case 'Boxes':
      for (let i = 0; i < 4; i++) {
        ctx.fillStyle = i < index ? '#2F6EBA' : '#C9D7EA';
        ctx.fillRect(x + 1 + (i % 2) * (size / 2), y + 1 + Math.floor(i / 2) * (size / 2), size / 2 - 1.5, size / 2 - 1.5);
      }
      break;
    default: {
      // Traffic lights and red-to-black circles.
      const fill = kind === 'RedToBlack' ? ['#1F1F1F', '#7F7F7F', '#E6A0A0', RED][index]! : n === 4 && kind === 'TrafficLights' ? ['#1F1F1F', RED, YELLOW, GREEN][index]! : color;
      if (kind === 'TrafficLights2') {
        ctx.fillStyle = '#3A3A3A';
        ctx.beginPath();
        ctx.arc(cx, cy, r, 0, Math.PI * 2);
        ctx.fill();
      }
      ctx.fillStyle = fill;
      ctx.beginPath();
      ctx.arc(cx, cy, kind === 'TrafficLights2' ? r * 0.7 : r, 0, Math.PI * 2);
      ctx.fill();
    }
  }
  ctx.restore();
}
