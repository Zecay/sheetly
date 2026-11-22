// The layer behind the app: gradient or picture, tiled texture, and an animated scene.
// Scenes draw to a small canvas at a capped frame rate from pre-rendered sprites, and never make the grid redraw.
import { Look, alpha, isDark, mix } from './theme';

type G = CanvasRenderingContext2D;
interface Pal {
  a: string;
  b: string;
  dark: boolean;
}
type Draw = (c: G, t: number, dt: number) => void;
type Scene = (w: number, h: number, p: Pal, n: number) => Draw;

const TAU = Math.PI * 2;
const rnd = (lo = 1, hi?: number) => (hi === undefined ? Math.random() * lo : lo + Math.random() * (hi - lo));

function sprite(w: number, h: number, paint: (g: G) => void): HTMLCanvasElement {
  const s = document.createElement('canvas');
  s.width = Math.max(1, Math.ceil(w));
  s.height = Math.max(1, Math.ceil(h));
  paint(s.getContext('2d')!);
  return s;
}
function glow(r: number, color: string): HTMLCanvasElement {
  return sprite(r * 2, r * 2, (g) => {
    const f = g.createRadialGradient(r, r, 0, r, r, r);
    f.addColorStop(0, alpha(color, 0.9));
    f.addColorStop(0.25, alpha(color, 0.4));
    f.addColorStop(1, alpha(color, 0));
    g.fillStyle = f;
    g.fillRect(0, 0, r * 2, r * 2);
  });
}
function starfield(w: number, h: number, top: string, bottom: string, count: number): HTMLCanvasElement {
  return sprite(w, h, (g) => {
    const f = g.createLinearGradient(0, 0, 0, h);
    f.addColorStop(0, top);
    f.addColorStop(1, bottom);
    g.fillStyle = f;
    g.fillRect(0, 0, w, h);
    for (let i = 0; i < count; i++) {
      g.fillStyle = `rgba(255,255,255,${rnd(0.15, 0.9).toFixed(2)})`;
      const s = Math.random() < 0.08 ? 1.6 : 0.9;
      g.fillRect(rnd(w), rnd(h), s, s);
    }
  });
}

// ---------------- scenes ----------------
const blackhole: Scene = (w, h, p, n) => {
  const S = Math.min(w, h);
  const R = S * 0.13;
  const D = Math.ceil(R * 4.3);
  const hot = mix(p.a, '#ffe2b0', 0.4);
  const cool = mix(p.a, p.b, 0.55);
  const back = sprite(w, h, (g) => {
    g.drawImage(starfield(w, h, '#04030c', '#020206', Math.round((w * h) / 2600)), 0, 0);
    const f = g.createRadialGradient(w / 2, h / 2, R, w / 2, h / 2, S * 0.85);
    f.addColorStop(0, alpha(p.a, 0.16));
    f.addColorStop(1, alpha(p.a, 0));
    g.fillStyle = f;
    g.fillRect(0, 0, w, h);
  });
  const streaks = (g: G, r0: number, r1: number, count: number, power: number) => {
    g.globalCompositeOperation = 'lighter';
    g.lineCap = 'round';
    for (let i = 0; i < count; i++) {
      const u = Math.random() ** power;
      const a0 = rnd(TAU);
      g.strokeStyle = alpha(u < 0.2 ? mix('#ffffff', hot, u / 0.2) : mix(hot, cool, (u - 0.2) / 0.8), (1 - u) * 0.26 + 0.025);
      g.lineWidth = rnd(0.5, 0.8 + S / 260);
      g.beginPath();
      g.arc(0, 0, r0 + u * (r1 - r0), a0, a0 + rnd(0.12, 0.9));
      g.stroke();
    }
  };
  const disk = sprite(D * 2, D * 2, (g) => {
    g.translate(D, D);
    streaks(g, R * 1.28, D * 0.98, 1500, 1.7);
  });
  const LR = Math.ceil(R * 1.75);
  const lens = sprite(LR * 2, LR * 2, (g) => {
    g.translate(LR, LR);
    streaks(g, R * 1.03, LR * 0.97, 520, 2.2);
  });
  const halo = glow(Math.ceil(R * 2.6), hot);
  const sparks = Array.from({ length: Math.round(70 * n) }, () => ({ r: rnd(1.35, 4), a: rnd(TAU), s: rnd(0.6, 1.3) }));
  const drawDisk = (c: G, t: number) => {
    c.save();
    c.scale(1, 0.3);
    c.rotate(t * 0.34);
    c.drawImage(disk, -D, -D);
    c.rotate(-t * 0.56);
    c.globalAlpha = 0.55;
    c.drawImage(disk, -D, -D);
    c.restore();
  };
  const drawSparks = (c: G, front: boolean) => {
    c.fillStyle = '#fff1d6';
    for (const k of sparks) {
      const sin = Math.sin(k.a);
      if (sin > 0 !== front) continue;
      c.fillRect(Math.cos(k.a) * k.r * R, sin * k.r * R * 0.3, 1.5, 1.5);
    }
  };
  return (c, t, dt) => {
    for (const k of sparks) {
      k.a += (dt * 1.5 * k.s) / k.r;
      k.r -= dt * 0.05 * k.s;
      if (k.r < 1.3) {
        k.r = rnd(3, 4.1);
        k.a = rnd(TAU);
      }
    }
    c.setTransform(1, 0, 0, 1, 0, 0);
    c.globalCompositeOperation = 'source-over';
    c.drawImage(back, 0, 0);
    c.translate(w / 2, h / 2);
    c.rotate(-0.2);
    const big = Math.max(w, h) * 1.2;
    c.globalCompositeOperation = 'lighter';
    c.globalAlpha = 0.45;
    c.drawImage(halo, -halo.width / 2, -halo.height / 2);
    c.globalAlpha = 1;
    // far half of the disk, then its image bent over the top of the shadow
    c.save();
    c.beginPath();
    c.rect(-big, -big, big * 2, big);
    c.clip();
    drawDisk(c, t);
    c.restore();
    drawSparks(c, false);
    c.save();
    c.rotate(t * 0.5);
    c.drawImage(lens, -LR, -LR);
    c.restore();
    c.globalCompositeOperation = 'source-over';
    c.fillStyle = '#000';
    c.beginPath();
    c.arc(0, 0, R, 0, TAU);
    c.fill();
    c.globalCompositeOperation = 'lighter';
    c.strokeStyle = alpha(hot, 0.85);
    c.lineWidth = Math.max(1, R * 0.028);
    c.beginPath();
    c.arc(0, 0, R * 1.01, 0, TAU);
    c.stroke();
    // near half passes in front
    c.save();
    c.beginPath();
    c.rect(-big, 0, big * 2, big);
    c.clip();
    drawDisk(c, t);
    c.restore();
    drawSparks(c, true);
    c.globalCompositeOperation = 'source-over';
    c.setTransform(1, 0, 0, 1, 0, 0);
  };
};

const warp: Scene = (w, h, p, n) => {
  const cx = w / 2;
  const cy = h / 2;
  const reset = (s: { x: number; y: number; z: number }, any = false) => {
    s.x = rnd(-w, w);
    s.y = rnd(-h, h);
    s.z = any ? rnd(0.05, 1) : 1;
  };
  const stars = Array.from({ length: Math.round(230 * n) }, () => {
    const s = { x: 0, y: 0, z: 1 };
    reset(s, true);
    return s;
  });
  const cols = [alpha(mix(p.a, '#ffffff', 0.3), 0.35), alpha(mix(p.a, '#ffffff', 0.55), 0.65), 'rgba(255,255,255,0.95)'];
  const bg = sprite(w, h, (g) => {
    g.fillStyle = '#03040c';
    g.fillRect(0, 0, w, h);
    const f = g.createRadialGradient(cx, cy, 0, cx, cy, Math.max(w, h) * 0.6);
    f.addColorStop(0, alpha(p.a, 0.2));
    f.addColorStop(1, alpha(p.a, 0));
    g.fillStyle = f;
    g.fillRect(0, 0, w, h);
  });
  return (c, _t, dt) => {
    c.drawImage(bg, 0, 0);
    c.lineWidth = 1;
    for (let b = 0; b < 3; b++) {
      c.strokeStyle = cols[b];
      c.beginPath();
      for (const s of stars) {
        const band = s.z > 0.6 ? 0 : s.z > 0.28 ? 1 : 2;
        if (band !== b) continue;
        const z0 = s.z + dt * 0.28;
        const x = cx + s.x / s.z / 2.2;
        const y = cy + s.y / s.z / 2.2;
        c.moveTo(x, y);
        c.lineTo(cx + s.x / z0 / 2.2, cy + s.y / z0 / 2.2 - 0);
      }
      c.stroke();
    }
    for (const s of stars) {
      s.z -= dt * 0.24;
      if (s.z < 0.04 || Math.abs(s.x / s.z / 2.2) > w || Math.abs(s.y / s.z / 2.2) > h) reset(s);
    }
  };
};

const aurora: Scene = (w, h, p) => {
  const bg = starfield(w, h, '#040a18', '#0a1026', Math.round((w * h) / 3500));
  const colors = [p.a, mix(p.a, p.b, 0.6), mix(p.a, '#7c4dff', 0.65)];
  const bands = colors.map((col, i) => {
    const y = h * (0.26 + i * 0.1);
    const tall = h * (0.34 - i * 0.04);
    const f = bg.getContext('2d')!.createLinearGradient(0, y - tall, 0, y + 30);
    f.addColorStop(0, alpha(col, 0));
    f.addColorStop(0.72, alpha(col, 0.2));
    f.addColorStop(1, alpha(col, 0));
    return { f, y, tall, phase: i * 1.9, k: 0.0035 + i * 0.0012 };
  });
  const step = Math.max(10, w / 44);
  return (c, t) => {
    c.globalCompositeOperation = 'source-over';
    c.drawImage(bg, 0, 0);
    c.globalCompositeOperation = 'lighter';
    for (const b of bands) {
      const yAt = (x: number) => b.y + Math.sin(x * b.k + t * 0.32 + b.phase) * h * 0.07 + Math.sin(x * b.k * 2.3 - t * 0.21) * h * 0.03;
      c.fillStyle = b.f;
      c.beginPath();
      c.moveTo(0, yAt(0) - b.tall);
      for (let x = step; x <= w + step; x += step) c.lineTo(x, yAt(x) - b.tall * (0.86 + 0.14 * Math.sin(x * 0.012 + t * 0.5)));
      for (let x = w + step; x >= 0; x -= step) c.lineTo(x, yAt(x) + 26);
      c.closePath();
      c.fill();
    }
    c.globalCompositeOperation = 'source-over';
  };
};

const waves: Scene = (w, h, p) => {
  const layers = Array.from({ length: 5 }, (_, i) => ({
    y: h * (0.5 + i * 0.1),
    amp: h * (0.04 - i * 0.004),
    k: 0.006 + i * 0.0016,
    v: 0.5 + i * 0.16,
    col: alpha(mix(p.a, p.dark ? '#02101f' : '#ffffff', 0.15 + i * 0.13), 0.3),
  }));
  const step = Math.max(8, w / 70);
  return (c, t) => {
    c.clearRect(0, 0, w, h);
    for (const l of layers) {
      c.fillStyle = l.col;
      c.beginPath();
      c.moveTo(0, h);
      for (let x = 0; x <= w + step; x += step) c.lineTo(x, l.y + Math.sin(x * l.k + t * l.v) * l.amp + Math.sin(x * l.k * 2.1 - t * l.v * 0.7) * l.amp * 0.4);
      c.lineTo(w + step, h);
      c.closePath();
      c.fill();
    }
  };
};

const fireflies: Scene = (w, h, p, n) => {
  const R = Math.max(7, Math.min(w, h) * 0.022);
  const spr = glow(R, mix(p.a, '#fff6a8', 0.3));
  const flies = Array.from({ length: Math.round(46 * n) }, () => ({ x: rnd(w), y: rnd(h), a: rnd(TAU), v: rnd(6, 20), ph: rnd(TAU), s: rnd(0.45, 1.1) }));
  return (c, t, dt) => {
    c.clearRect(0, 0, w, h);
    c.globalCompositeOperation = 'lighter';
    for (const f of flies) {
      f.a += Math.sin(t * 0.6 + f.ph) * dt * 1.4;
      f.x = (f.x + Math.cos(f.a) * f.v * dt + w) % w;
      f.y = (f.y + Math.sin(f.a) * f.v * dt + h) % h;
      c.globalAlpha = 0.25 + 0.75 * Math.max(0, Math.sin(t * 1.3 + f.ph)) ** 2;
      const d = R * 2 * f.s;
      c.drawImage(spr, f.x - d / 2, f.y - d / 2, d, d);
    }
    c.globalAlpha = 1;
    c.globalCompositeOperation = 'source-over';
  };
};

const snow: Scene = (w, h, p, n) => {
  const col = p.dark ? '#ffffff' : '#ffffff';
  const spr = glow(5, col);
  const flakes = Array.from({ length: Math.round(170 * n) }, () => {
    const z = rnd(0.35, 1);
    return { x: rnd(w), y: rnd(h), z, ph: rnd(TAU) };
  });
  return (c, t, dt) => {
    c.clearRect(0, 0, w, h);
    for (const f of flakes) {
      f.y += f.z * 34 * dt;
      f.x += Math.sin(t * 0.7 + f.ph) * 12 * f.z * dt;
      if (f.y > h + 6) {
        f.y = -6;
        f.x = rnd(w);
      }
      if (f.x > w + 6) f.x = -6;
      else if (f.x < -6) f.x = w + 6;
      const d = 7 * f.z;
      c.globalAlpha = 0.45 + f.z * 0.55;
      c.drawImage(spr, f.x - d / 2, f.y - d / 2, d, d);
    }
    c.globalAlpha = 1;
  };
};

const rain: Scene = (w, h, p, n) => {
  const drops = Array.from({ length: Math.round(220 * n) }, () => {
    const z = rnd(0.4, 1);
    return { x: rnd(-40, w + 40), y: rnd(h), z, b: z > 0.72 ? 1 : 0 };
  });
  const cols = [alpha(mix(p.a, '#ffffff', 0.3), 0.22), alpha(mix(p.a, '#ffffff', 0.55), 0.5)];
  return (c, _t, dt) => {
    c.clearRect(0, 0, w, h);
    c.lineWidth = 1;
    for (let b = 0; b < 2; b++) {
      c.strokeStyle = cols[b];
      c.beginPath();
      for (const d of drops) {
        if (d.b !== b) continue;
        d.y += d.z * 520 * dt;
        d.x -= d.z * 70 * dt;
        if (d.y > h + 20) {
          d.y = -20;
          d.x = rnd(-40, w + 60);
        }
        c.moveTo(d.x, d.y);
        c.lineTo(d.x - d.z * 2.4, d.y + d.z * 17);
      }
      c.stroke();
    }
  };
};

const constellation: Scene = (w, h, p, n) => {
  const pts = Array.from({ length: Math.round(64 * n) }, () => ({ x: rnd(w), y: rnd(h), vx: rnd(-9, 9), vy: rnd(-9, 9) }));
  const link = Math.min(w, h) * 0.2;
  const link2 = link * link;
  const line = mix(p.a, '#ffffff', 0.25);
  const cols = [alpha(line, 0.08), alpha(line, 0.2), alpha(line, 0.38)];
  return (c, _t, dt) => {
    c.clearRect(0, 0, w, h);
    for (const a of pts) {
      a.x += a.vx * dt;
      a.y += a.vy * dt;
      if (a.x < 0 || a.x > w) a.vx = -a.vx;
      if (a.y < 0 || a.y > h) a.vy = -a.vy;
    }
    c.lineWidth = 1;
    for (let b = 0; b < 3; b++) {
      c.strokeStyle = cols[b];
      c.beginPath();
      for (let i = 0; i < pts.length; i++) {
        for (let j = i + 1; j < pts.length; j++) {
          const dx = pts[i].x - pts[j].x;
          const dy = pts[i].y - pts[j].y;
          const d = dx * dx + dy * dy;
          if (d > link2) continue;
          if (Math.min(2, Math.floor((1 - d / link2) * 3)) !== b) continue;
          c.moveTo(pts[i].x, pts[i].y);
          c.lineTo(pts[j].x, pts[j].y);
        }
      }
      c.stroke();
    }
    c.fillStyle = mix(p.a, '#ffffff', 0.55);
    for (const a of pts) c.fillRect(a.x - 1.2, a.y - 1.2, 2.4, 2.4);
  };
};

const lava: Scene = (w, h, p, n) => {
  const S = Math.max(w, h);
  const sprites = [glow(Math.ceil(S * 0.2), p.a), glow(Math.ceil(S * 0.2), mix(p.a, p.b, 0.55)), glow(Math.ceil(S * 0.2), mix(p.a, '#ffc53d', 0.6))];
  const blobs = Array.from({ length: Math.round(9 * n) + 2 }, (_, i) => ({ x: rnd(w), ph: rnd(TAU), v: rnd(0.1, 0.24), r: rnd(0.45, 1), spr: sprites[i % 3] }));
  return (c, t) => {
    c.clearRect(0, 0, w, h);
    c.globalCompositeOperation = 'lighter';
    c.globalAlpha = 0.62;
    for (const b of blobs) {
      const d = S * 0.4 * b.r * (1 + 0.14 * Math.sin(t * 0.5 + b.ph));
      const y = h / 2 + Math.sin(t * b.v + b.ph) * h * 0.62;
      const x = b.x + Math.sin(t * b.v * 0.6 + b.ph * 2) * w * 0.06;
      c.drawImage(b.spr, x - d / 2, y - d / 2, d, d);
    }
    c.globalAlpha = 1;
    c.globalCompositeOperation = 'source-over';
  };
};

const retro: Scene = (w, h, p) => {
  const hz = Math.round(h * 0.56);
  const R = Math.min(w, h) * 0.24;
  const sky = sprite(w, h, (g) => {
    g.drawImage(starfield(w, hz, '#10042a', '#3a0d58', Math.round((w * hz) / 3200)), 0, 0);
    const haze = g.createLinearGradient(0, hz - h * 0.2, 0, hz);
    haze.addColorStop(0, alpha(p.a, 0));
    haze.addColorStop(1, alpha(p.a, 0.5));
    g.fillStyle = haze;
    g.fillRect(0, hz - h * 0.2, w, h * 0.2);
    const sun = g.createLinearGradient(0, hz - R * 1.3, 0, hz);
    sun.addColorStop(0, '#ffe66d');
    sun.addColorStop(0.55, '#ff8a5c');
    sun.addColorStop(1, p.a);
    g.fillStyle = sun;
    g.beginPath();
    g.arc(w / 2, hz - R * 0.34, R, 0, TAU);
    g.fill();
    // bands cut through the lower part of the sun
    g.globalCompositeOperation = 'destination-out';
    for (let i = 0; i < 7; i++) g.fillRect(w / 2 - R * 1.1, hz - R * 0.72 + i * R * 0.13, R * 2.2, 1 + i * R * 0.012);
    g.globalCompositeOperation = 'destination-over';
    g.fillStyle = '#25094a';
    g.fillRect(0, 0, w, hz);
    g.globalCompositeOperation = 'source-over';
    const floor = g.createLinearGradient(0, hz, 0, h);
    floor.addColorStop(0, '#1b0636');
    floor.addColorStop(1, '#07020f');
    g.fillStyle = floor;
    g.fillRect(0, hz, w, h - hz);
  });
  const line = alpha(mix(p.a, '#ff8ad1', 0.3), 0.75);
  const verticals = Math.ceil(w / 34);
  return (c, t) => {
    c.drawImage(sky, 0, 0);
    c.strokeStyle = line;
    c.lineWidth = 1;
    c.beginPath();
    for (let i = -verticals; i <= verticals; i++) {
      c.moveTo(w / 2 + (i * w) / verticals / 7, hz);
      c.lineTo(w / 2 + (i * w) / verticals * 1.6, h);
    }
    const f = (t * 0.32) % 1;
    for (let i = 0; i < 13; i++) {
      const k = (i + f) / 13;
      const y = hz + (h - hz) * k * k * k;
      c.moveTo(0, y);
      c.lineTo(w, y);
    }
    c.stroke();
  };
};

export const SCENES: [string, string, Scene][] = [
  ['blackhole', 'Black Hole', blackhole],
  ['warp', 'Warp Speed', warp],
  ['aurora', 'Aurora', aurora],
  ['retro', 'Retro Sun', retro],
  ['waves', 'Waves', waves],
  ['fireflies', 'Fireflies', fireflies],
  ['rain', 'Rain', rain],
  ['snow', 'Snowfall', snow],
  ['lava', 'Lava Lamp', lava],
  ['constellation', 'Constellation', constellation],
];
const sceneOf = (id: string) => SCENES.find((s) => s[0] === id)?.[2] ?? null;
const palOf = (l: Look): Pal => ({ a: l.c.accent, b: l.wall.mode === 'gradient' ? mix(l.wall.a, l.c.bar, 0.5) : l.c.bar, dark: isDark(l) });

/** One still frame for a picker card. */
export function previewScene(cv: HTMLCanvasElement, id: string, l: Look): void {
  const s = sceneOf(id);
  const c = cv.getContext('2d');
  if (!s || !c) return;
  c.fillStyle = l.wall.mode === 'gradient' ? mix(l.wall.a, l.wall.b, 0.5) : sceneBase(l);
  c.fillRect(0, 0, cv.width, cv.height);
  // Scenes clear their own canvas, so the frame is drawn separately and laid over the backdrop.
  const d = s(cv.width, cv.height, palOf(l), 0.7);
  c.drawImage(sprite(cv.width, cv.height, (g) => d(g, 3, 0.02)), 0, 0);
}
/** Glowing and white scenes need a deep backdrop: on a plain white theme they would be invisible. */
const sceneBase = (l: Look) => mix(l.c.accent, '#0a0a14', 0.82);

// ---------------- textures ----------------
type Tile = (g: G, s: number) => void;
function seeded(seed: number) {
  return () => {
    seed = (seed * 16807) % 2147483647;
    return seed / 2147483647;
  };
}
const speckle = (g: G, s: number, count: number, size: number, stretch = 1) => {
  const r = seeded(7919);
  for (let i = 0; i < count; i++) {
    g.globalAlpha = 0.08 + r() * 0.3;
    g.fillRect(r() * s, r() * s, size * stretch, size);
  }
  g.globalAlpha = 1;
};
export const TEXTURES: [string, string, number, Tile][] = [
  ['dots', 'Dots', 24, (g) => { g.beginPath(); g.arc(12, 12, 1.6, 0, TAU); g.fill(); }],
  ['grid', 'Blueprint', 32, (g, s) => { g.strokeRect(0.5, 0.5, s, s); }],
  ['diagonal', 'Pinstripe', 16, (g, s) => { g.beginPath(); g.moveTo(-2, s + 2); g.lineTo(s + 2, -2); g.moveTo(-2, 2); g.lineTo(2, -2); g.moveTo(s - 2, s + 2); g.lineTo(s + 2, s - 2); g.stroke(); }],
  ['cross', 'Crosshatch', 20, (g, s) => { g.globalAlpha = 0.7; g.beginPath(); g.moveTo(0, 0); g.lineTo(s, s); g.moveTo(s, 0); g.lineTo(0, s); g.stroke(); g.globalAlpha = 1; }],
  ['carbon', 'Carbon', 12, (g) => { g.globalAlpha = 0.55; g.fillRect(0, 0, 6, 6); g.fillRect(6, 6, 6, 6); g.globalAlpha = 0.2; g.fillRect(6, 0, 6, 6); g.fillRect(0, 6, 6, 6); g.globalAlpha = 1; }],
  ['paper', 'Paper', 160, (g, s) => speckle(g, s, 1500, 1)],
  ['grain', 'Film Grain', 128, (g, s) => speckle(g, s, 3200, 1)],
  ['linen', 'Linen', 96, (g, s) => { speckle(g, s, 420, 1, 14); g.save(); g.translate(s, 0); g.rotate(Math.PI / 2); speckle(g, s, 420, 1, 14); g.restore(); }],
  ['brushed', 'Brushed Metal', 200, (g, s) => speckle(g, s, 900, 1, 46)],
  ['hex', 'Honeycomb', 42, (g) => { g.beginPath(); g.moveTo(0, 12); g.lineTo(10.5, 6); g.lineTo(21, 12); g.lineTo(31.5, 6); g.lineTo(42, 12); g.moveTo(10.5, 6); g.lineTo(10.5, 0); g.moveTo(31.5, 6); g.lineTo(31.5, 0); g.moveTo(0, 12); g.lineTo(0, 30); g.moveTo(21, 12); g.lineTo(21, 30); g.moveTo(42, 12); g.lineTo(42, 30); g.moveTo(0, 30); g.lineTo(10.5, 36); g.lineTo(21, 30); g.lineTo(31.5, 36); g.lineTo(42, 30); g.moveTo(10.5, 36); g.lineTo(10.5, 42); g.moveTo(31.5, 36); g.lineTo(31.5, 42); g.stroke(); }],
  ['scales', 'Scales', 40, (g) => { g.beginPath(); for (const [x, y] of [[0, 0], [40, 0], [20, 20], [0, 40], [40, 40]]) { g.moveTo(x + 20, y); g.arc(x, y, 20, 0, Math.PI); } g.stroke(); }],
  ['circuit', 'Circuit', 64, (g) => { g.beginPath(); g.moveTo(0, 16); g.lineTo(20, 16); g.lineTo(32, 28); g.lineTo(32, 64); g.moveTo(32, 0); g.lineTo(32, 8); g.lineTo(46, 22); g.lineTo(64, 22); g.moveTo(0, 48); g.lineTo(14, 48); g.moveTo(48, 64); g.lineTo(48, 44); g.lineTo(64, 44); g.stroke(); for (const [x, y] of [[14, 48], [48, 44], [20, 16]]) { g.beginPath(); g.arc(x, y, 2.4, 0, TAU); g.fill(); } }],
];
const texCache = new Map<string, string>();
export function textureUrl(id: string, dark: boolean): string | null {
  const t = TEXTURES.find((x) => x[0] === id);
  if (!t) return null;
  const key = id + (dark ? ':d' : ':l');
  let url = texCache.get(key);
  if (!url) {
    const ink = dark ? '#ffffff' : '#000000';
    url = sprite(t[2], t[2], (g) => {
      g.fillStyle = ink;
      g.strokeStyle = ink;
      g.lineWidth = 1;
      t[3](g, t[2]);
    }).toDataURL('image/png');
    texCache.set(key, url);
  }
  return url;
}

// ---------------- the live layer ----------------
const QUALITY = { eco: { scale: 0.5, fps: 20 }, balanced: { scale: 0.67, fps: 30 }, smooth: { scale: 1, fps: 60 } };

export class Wallpaper {
  private root = document.getElementById('wallpaper')!;
  private base = document.getElementById('wp-base')!;
  private tex = document.getElementById('wp-tex')!;
  private cv = document.getElementById('wp-anim') as HTMLCanvasElement;
  private ctx = this.cv.getContext('2d')!;
  private draw: Draw | null = null;
  private look: Look | null = null;
  private key = '';
  private raf = 0;
  private last = 0;
  private t = 0;
  private resizeTimer = 0;
  private still = window.matchMedia('(prefers-reduced-motion: reduce)');

  constructor() {
    const sync = () => this.sync();
    document.addEventListener('visibilitychange', sync);
    window.addEventListener('focus', sync);
    window.addEventListener('blur', sync);
    this.still.addEventListener('change', sync);
    window.addEventListener('resize', () => {
      clearTimeout(this.resizeTimer);
      this.resizeTimer = window.setTimeout(() => this.rebuild(), 180);
    });
  }

  get reducedMotion(): boolean {
    return this.still.matches;
  }

  apply(l: Look, image: string | null): void {
    this.look = l;
    const active = l.wall.mode !== 'none' || l.texture.id !== 'none' || l.anim.id !== 'none';
    this.root.hidden = !active;
    const b = this.base.style;
    if (l.wall.mode === 'image' && image) {
      b.background = `center / cover no-repeat url("${image}")`;
    } else if (l.wall.mode === 'gradient') {
      b.background = l.wall.radial ? `radial-gradient(circle at 50% 40%, ${l.wall.a}, ${l.wall.b})` : `linear-gradient(${l.wall.angle}deg, ${l.wall.a}, ${l.wall.b})`;
    } else b.background = l.anim.id !== 'none' && sceneOf(l.anim.id) ? sceneBase(l) : l.c.sheet;
    const url = textureUrl(l.texture.id, isDark(l));
    const tile = TEXTURES.find((x) => x[0] === l.texture.id)?.[2] ?? 32;
    this.tex.style.backgroundImage = url ? `url("${url}")` : 'none';
    this.tex.style.backgroundSize = `${Math.round((tile * l.texture.scale) / 100)}px`;
    this.tex.style.opacity = String((l.texture.strength / 100) * 0.55);
    const key = [l.anim.id, l.anim.amount, l.anim.quality, l.c.accent, l.c.bar, l.wall.a, l.wall.b, l.wall.mode].join('|');
    if (key !== this.key) {
      this.key = key;
      this.rebuild();
    } else this.sync();
  }

  private rebuild(): void {
    const l = this.look;
    const scene = l && !this.root.hidden ? sceneOf(l.anim.id) : null;
    this.cv.hidden = !scene;
    if (!l || !scene) {
      this.draw = null;
      this.sync();
      return;
    }
    const q = QUALITY[l.anim.quality];
    const k = Math.min(q.scale, 1600 / Math.max(1, window.innerWidth));
    const w = Math.max(160, Math.round(window.innerWidth * k));
    const h = Math.max(100, Math.round(window.innerHeight * k));
    this.cv.width = w;
    this.cv.height = h;
    this.draw = scene(w, h, palOf(l), l.anim.amount / 100);
    this.frame(0.016);
    this.sync();
  }

  private frame(dt: number): void {
    if (!this.draw || !this.look) return;
    this.t += dt;
    this.draw(this.ctx, this.t, dt);
  }

  /** Runs the loop only while someone can see it. */
  private sync(): void {
    const l = this.look;
    const run = !!this.draw && !!l && !document.hidden && !this.still.matches && !(l.anim.pauseUnfocused && !document.hasFocus());
    if (run && !this.raf) {
      this.last = performance.now();
      this.raf = requestAnimationFrame(this.tick);
    } else if (!run && this.raf) {
      cancelAnimationFrame(this.raf);
      this.raf = 0;
    }
  }

  private tick = (now: number): void => {
    this.raf = requestAnimationFrame(this.tick);
    const l = this.look!;
    const gap = now - this.last;
    if (gap < 1000 / QUALITY[l.anim.quality].fps - 2) return;
    this.last = now;
    if (!document.getElementById('backstage')!.hidden) return;
    this.frame((Math.min(gap, 100) / 1000) * (l.anim.speed / 100));
  };
}
