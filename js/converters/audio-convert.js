import { getLame } from '../lib/loader.js';
import { baseName, tick } from '../lib/util.js';

const MAX_FILE_BYTES = 150 * 1024 * 1024; // a decoded track takes ~10x its compressed size in memory
const MAX_SECONDS = 3 * 3600;
const RATE = 44100; // every track is resampled to 44.1 kHz, which every MP3 player supports

async function decode(file, report) {
  const Ctx = window.OfflineAudioContext || window.webkitOfflineAudioContext;
  if (!Ctx) throw new Error('This browser can’t read audio files. Try a recent Chrome, Edge, Firefox or Safari.');
  const bytes = await file.arrayBuffer();
  report(0.1, 'Reading the audio');
  const ctx = new Ctx(2, 1, RATE);
  try {
    return await ctx.decodeAudioData(bytes);
  } catch (e) {
    if (e && e.name === 'NotSupportedError') throw new Error('This browser can’t play this kind of audio file, so it can’t convert it. Try Chrome or Edge, or a different file.');
    throw new Error('This audio file couldn’t be read. It may be damaged, copy-protected, or in a format this browser doesn’t support.');
  }
}

// Returns one or two channels of samples. Surround tracks are mixed down to stereo.
function channelsOf(buffer) {
  const n = buffer.numberOfChannels;
  if (n <= 2) return Array.from({ length: n }, (_, i) => buffer.getChannelData(i));
  const len = buffer.length;
  const left = new Float32Array(len);
  const right = new Float32Array(len);
  let l = 0;
  let r = 0;
  for (let c = 0; c < n; c++) {
    const data = buffer.getChannelData(c);
    const dest = c % 2 === 0 ? left : right;
    if (c % 2 === 0) l++; else r++;
    for (let i = 0; i < len; i++) dest[i] += data[i];
  }
  for (let i = 0; i < len; i++) { left[i] /= l; right[i] /= (r || 1); }
  return [left, right];
}

const toInt16 = (src, from, to) => {
  const out = new Int16Array(to - from);
  for (let i = from, j = 0; i < to; i++, j++) {
    const v = src[i];
    out[j] = v !== v ? 0 : v >= 1 ? 32767 : v <= -1 ? -32768 : Math.round(v * 32767);
  }
  return out;
};

async function encodeMp3(channels, kbps, report) {
  const lamejs = await getLame();
  const encoder = new lamejs.Mp3Encoder(channels.length, RATE, kbps);
  const total = channels[0].length;
  const block = 1152 * 40;
  const parts = [];
  for (let pos = 0; pos < total; pos += block) {
    const end = Math.min(total, pos + block);
    const left = toInt16(channels[0], pos, end);
    const chunk = channels.length === 2
      ? encoder.encodeBuffer(left, toInt16(channels[1], pos, end))
      : encoder.encodeBuffer(left);
    if (chunk.length) parts.push(chunk);
    if (((pos / block) & 7) === 0) {
      report(0.2 + 0.78 * (pos / total), 'Encoding MP3');
      await tick();
    }
  }
  const tail = encoder.flush();
  if (tail.length) parts.push(tail);
  return new Blob(parts, { type: 'audio/mpeg' });
}

function encodeWav(channels) {
  const n = channels.length;
  const frames = channels[0].length;
  const dataSize = frames * n * 2;
  if (dataSize > 0xfffffff0 - 36) throw new Error('This track is too long to save as WAV (the format is limited to 4 GB).');
  const out = new ArrayBuffer(44 + dataSize);
  const dv = new DataView(out);
  const str = (o, s) => { for (let i = 0; i < s.length; i++) dv.setUint8(o + i, s.charCodeAt(i)); };
  str(0, 'RIFF'); dv.setUint32(4, 36 + dataSize, true); str(8, 'WAVE');
  str(12, 'fmt '); dv.setUint32(16, 16, true); dv.setUint16(20, 1, true); dv.setUint16(22, n, true);
  dv.setUint32(24, RATE, true); dv.setUint32(28, RATE * n * 2, true); dv.setUint16(32, n * 2, true); dv.setUint16(34, 16, true);
  str(36, 'data'); dv.setUint32(40, dataSize, true);
  const pcm = new Int16Array(out, 44);
  for (let c = 0; c < n; c++) {
    const src = channels[c];
    for (let i = 0, j = c; i < frames; i++, j += n) {
      const v = src[i];
      pcm[j] = v !== v ? 0 : v >= 1 ? 32767 : v <= -1 ? -32768 : Math.round(v * 32767);
    }
  }
  return new Blob([out], { type: 'audio/wav' });
}

export async function convert(file, { report, options }) {
  if (!file.size) throw new Error('This file is empty.');
  if (file.size > MAX_FILE_BYTES) throw new Error(`This file is larger than ${Math.round(MAX_FILE_BYTES / 1048576)} MB, which is more than a browser can safely convert.`);
  const toWav = options.format === 'wav';
  const kbps = [128, 192, 256, 320].includes(Number(options.bitrate)) ? Number(options.bitrate) : 192;

  report(0.03, 'Opening the file');
  const buffer = await decode(file, report);
  if (!buffer.length || !buffer.duration) throw new Error('This audio file has no sound in it.');
  if (buffer.duration > MAX_SECONDS) throw new Error('This track is longer than 3 hours, which is more than a browser can safely convert.');

  const warnings = [];
  if (buffer.numberOfChannels > 2) warnings.push(`This track has ${buffer.numberOfChannels} channels, so it was mixed down to stereo.`);
  const channels = channelsOf(buffer);

  const mins = Math.floor(buffer.duration / 60);
  const secs = String(Math.round(buffer.duration % 60)).padStart(2, '0');
  const detail = `${mins}:${secs}${toWav ? ' · 16-bit' : ` · ${kbps} kbps`} · ${channels.length === 1 ? 'mono' : 'stereo'}`;

  if (toWav) {
    report(0.5, 'Writing WAV');
    await tick();
    return { blob: encodeWav(channels), name: `${baseName(file.name)}${/\.wav$/i.test(file.name) ? '-converted' : ''}.wav`, detail, warnings };
  }
  const blob = await encodeMp3(channels, kbps, report);
  return { blob, name: `${baseName(file.name)}${/\.mp3$/i.test(file.name) ? '-converted' : ''}.mp3`, detail, warnings };
}
