// browser.js — Playwright browser lifecycle and page interaction

import { chromium } from 'playwright';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { extname, join } from 'node:path';
import { homedir } from 'node:os';

let browserInstance = null;
let page = null;
let downloads = [];
let cdpSession = null;
let deviceNameFilters = [];
let discoveredDevices = [];
let machineNames = [];
let lastImageInfo = null;

// Read pixel/vector dimensions directly from the file.
// mods's moduleOutput event doesn't carry width/height, so this is the
// authoritative source for lastImageInfo after a file load.
async function readFileDimensions(filePath) {
  const ext = extname(filePath).toLowerCase();
  if (ext !== '.png' && ext !== '.svg') return null;
  const buf = await readFile(filePath);
  if (ext === '.png') {
    // PNG IHDR: width at byte 16, height at byte 20 (big-endian uint32)
    return { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) };
  }
  const svg = buf.toString('utf-8');
  const vb = svg.match(/viewBox=["'][\d.]+\s+[\d.]+\s+([\d.]+)\s+([\d.]+)["']/);
  if (vb) return { width: parseFloat(vb[1]), height: parseFloat(vb[2]) };
  const w = svg.match(/width=["']([\d.]+)/);
  const h = svg.match(/height=["']([\d.]+)/);
  if (w && h) return { width: parseFloat(w[1]), height: parseFloat(h[1]) };
  return null;
}

async function handleDevicePrompt(event) {
  const { id, devices } = event;
  console.error(`[mops] Device prompt: ${devices.map(d => d.name).join(', ')}`);

  for (const d of devices) {
    if (!discoveredDevices.some(dd => dd.name === d.name)) {
      discoveredDevices.push({ name: d.name, discoveredAt: Date.now() });
    }
  }

  // 1. Try exact deviceName filters first
  let match = devices.find(d =>
    deviceNameFilters.some(filter => d.name.toLowerCase().includes(filter.toLowerCase()))
  );

  // 2. Fuzzy match against profile machine names
  if (!match && machineNames.length > 0) {
    match = devices.find(d => {
      const dLower = d.name.toLowerCase();
      return machineNames.some(mName => {
        const keywords = mName.toLowerCase().split(/\s+/).filter(w => w.length > 2);
        const matched = keywords.filter(kw => dLower.includes(kw));
        return matched.length >= 2;
      });
    });
  }

  if (match) {
    console.error(`[mops] Auto-selecting device: ${match.name}`);
    if (!deviceNameFilters.some(f => f.toLowerCase() === match.name.toLowerCase())) {
      deviceNameFilters.push(match.name);
    }
    try {
      await cdpSession.send('DeviceAccess.selectPrompt', { id, deviceId: match.id });
    } catch (err) {
      console.error(`[mops] selectPrompt failed: ${err.message}`);
    }
  } else {
    console.error(`[mops] No auto-match. Discovered: ${devices.map(d => d.name).join(', ')}`);
  }
}

async function setupCdpDeviceAccess() {
  if (cdpSession) {
    try { await cdpSession.detach(); } catch {}
  }
  cdpSession = await page.context().newCDPSession(page);
  await cdpSession.send('DeviceAccess.enable');
  cdpSession.on('DeviceAccess.deviceRequestPrompted', handleDevicePrompt);
}

// Repair Chrome's crash markers if a prior session died hard (SIGKILL, etc).
// The restore-session bubble reads these fields at startup; clearing them is
// the real fix — the --disable-session-crashed-bubble flag is just belt.
async function markProfileCleanExit(userDataDir) {
  const prefsPath = join(userDataDir, 'Default', 'Preferences');
  try {
    const raw = await readFile(prefsPath, 'utf-8');
    const prefs = JSON.parse(raw);
    const profile = prefs.profile || (prefs.profile = {});
    if (profile.exit_type === 'Normal' && profile.exited_cleanly === true) return;
    profile.exit_type = 'Normal';
    profile.exited_cleanly = true;
    await writeFile(prefsPath, JSON.stringify(prefs));
  } catch { /* missing/malformed — fresh profile or we'll let Chrome recreate */ }
}

export async function launch(modsUrl, headless = false) {
  // Persistent profile so WebUSB/WebSerial device grants survive across sessions
  const userDataDir = join(homedir(), '.mops', 'chrome-data');
  await mkdir(userDataDir, { recursive: true });
  await markProfileCleanExit(userDataDir);
  const context = await chromium.launchPersistentContext(userDataDir, {
    headless, channel: 'chrome', acceptDownloads: true,
    args: ['--disable-session-crashed-bubble']
  });
  browserInstance = context;
  page = context.pages()[0] || await context.newPage();

  // Intercept downloads
  page.on('download', async (download) => {
    const path = await download.path();
    const content = path ? await readFile(path) : null;
    downloads.push({
      suggestedFilename: download.suggestedFilename(),
      content,
      timestamp: Date.now()
    });
  });

  await setupCdpDeviceAccess();

  await page.goto(modsUrl, { waitUntil: 'load' });
  await page.waitForFunction(() => typeof window.mods_prog_load === 'function', { timeout: 15000 });

  return page;
}

export function setDeviceFilters(filters, names) {
  deviceNameFilters = filters;
  machineNames = names;
}

export function getDiscoveredDevices() {
  return discoveredDevices;
}

export async function getGrantedDevices() {
  if (!page) return [];
  try {
    return await page.evaluate(async () => {
      const devices = [];
      if (navigator.usb && navigator.usb.getDevices) {
        const usbDevices = await navigator.usb.getDevices();
        for (const d of usbDevices) {
          devices.push({
            name: [d.manufacturerName, d.productName].filter(Boolean).join(' ') || 'Unknown USB device',
            vendorId: d.vendorId,
            productId: d.productId,
            serialNumber: d.serialNumber || null,
            type: 'usb'
          });
        }
      }
      return devices;
    });
  } catch {
    return [];
  }
}

export async function loadProgram(modsUrl, programPath, srcUrl) {
  if (!page) throw new Error('Browser not launched');
  lastImageInfo = null;
  const encodedPath = programPath.split('/').map(encodeURIComponent).join('/');
  let url = `${modsUrl}/?program=${encodedPath}`;
  if (srcUrl) url += `&src=${encodeURIComponent(srcUrl)}`;
  await page.goto(url, { waitUntil: 'load' });
  // Recreate CDP session after navigation (session state is unreliable after page.goto)
  await setupCdpDeviceAccess();
  await page.waitForFunction(() => typeof window.mods_prog_load === 'function', { timeout: 15000 });
  await page.waitForFunction(() => {
    const modules = document.getElementById('modules');
    return modules && modules.childNodes.length > 0;
  }, { timeout: 15000 });
  if (srcUrl) {
    // Wait for the src file to be processed
    await waitForProcessingSignal({ timeout: 10000 });
  }
}

export async function postMessageFile(filePath) {
  if (!page) throw new Error('Browser not launched');
  const ext = extname(filePath).toLowerCase();
  const fileData = await readFile(filePath);

  if (ext !== '.png' && ext !== '.svg') {
    return { error: `postMessage not supported for ${ext} files. Use setModuleFile for this type.` };
  }

  // Send file and wait for processing signal:
  // - New mods: moduleOutput event (with data like imageInfo)
  // - Old mods: 'ready' acknowledgment
  const msgType = ext === '.png' ? 'png' : 'svg';
  const payload = ext === '.png' ? fileData.toString('base64') : fileData.toString('utf-8');
  const outputEvent = await page.evaluate(({ type, data }) => {
    return new Promise(resolve => {
      const handler = (e) => {
        if (e.data && e.data.type === 'moduleOutput') {
          window.removeEventListener('message', handler);
          resolve(e.data);
        } else if (e.data === 'ready') {
          window.removeEventListener('message', handler);
          resolve('ready');
        }
      };
      window.addEventListener('message', handler);
      setTimeout(() => { window.removeEventListener('message', handler); resolve(null); }, 10000);
      window.postMessage({ type, data }, '*');
    });
  }, { type: msgType, data: payload });

  const result = { success: true, file: filePath, method: 'postMessage' };
  if (outputEvent && outputEvent !== 'ready') {
    result.module = outputEvent.module;
    result.output = outputEvent.output;
    if (outputEvent.data) result.moduleData = outputEvent.data;
  }
  const fileDims = await readFileDimensions(filePath);
  if (fileDims) {
    const mod = (outputEvent && outputEvent !== 'ready') ? outputEvent.module : null;
    const eventData = (outputEvent && outputEvent !== 'ready' && outputEvent.data) ? outputEvent.data : {};
    lastImageInfo = {
      ...eventData,
      width: fileDims.width,
      height: fileDims.height,
      moduleId: mod?.id ?? null,
      moduleName: mod?.name ?? null,
    };
  }
  return result;
}

export async function setModuleFile(moduleId, filePath) {
  if (!page) throw new Error('Browser not launched');
  const input = page.locator(`[id="${moduleId}"] input[type="file"]`);
  const count = await input.count();
  if (count === 0) {
    return { error: `No file input found in module ${moduleId}` };
  }
  // Wait for processing signal (moduleOutput on new mods, 'ready' on old, timeout as last resort)
  const eventPromise = page.evaluate((timeout) => {
    return new Promise(resolve => {
      const handler = (e) => {
        if (e.data && e.data.type === 'moduleOutput') {
          window.removeEventListener('message', handler);
          resolve(e.data);
        } else if (e.data === 'ready') {
          window.removeEventListener('message', handler);
          resolve('ready');
        }
      };
      window.addEventListener('message', handler);
      setTimeout(() => { window.removeEventListener('message', handler); resolve(null); }, timeout);
    });
  }, 10000);
  await input.setInputFiles(filePath);
  const outputEvent = await eventPromise;
  const result = { success: true, file: filePath, method: 'fileInput' };
  if (outputEvent && outputEvent !== 'ready') {
    result.module = outputEvent.module;
    result.output = outputEvent.output;
    if (outputEvent.data) result.moduleData = outputEvent.data;
  }
  const fileDims = await readFileDimensions(filePath);
  if (fileDims) {
    const mod = (outputEvent && outputEvent !== 'ready') ? outputEvent.module : null;
    const eventData = (outputEvent && outputEvent !== 'ready' && outputEvent.data) ? outputEvent.data : {};
    lastImageInfo = {
      ...eventData,
      width: fileDims.width,
      height: fileDims.height,
      moduleId: mod?.id ?? null,
      moduleName: mod?.name ?? null,
    };
  }
  return result;
}

export async function getProgramState() {
  if (!page) throw new Error('Browser not launched');
  return page.evaluate(() => {
    const modulesContainer = document.getElementById('modules');
    if (!modulesContainer) return [];

    const connections = {};
    const svg = document.getElementById('svg');
    if (svg) {
      const linksGroup = svg.getElementById('links');
      if (linksGroup) {
        for (let l = 0; l < linksGroup.childNodes.length; l++) {
          const link = linksGroup.childNodes[l];
          if (!link.id) continue;
          try {
            const linkData = JSON.parse(link.id);
            const source = JSON.parse(linkData.source);
            const dest = JSON.parse(linkData.dest);
            if (!connections[source.id]) connections[source.id] = { inputs: [], outputs: [] };
            if (!connections[dest.id]) connections[dest.id] = { inputs: [], outputs: [] };
            const srcMod = document.getElementById(source.id);
            const destMod = document.getElementById(dest.id);
            const srcName = srcMod ? srcMod.dataset.name : source.id;
            const destName = destMod ? destMod.dataset.name : dest.id;
            connections[source.id].outputs.push({ to: destName, toId: dest.id, port: source.name + ' → ' + dest.name });
            connections[dest.id].inputs.push({ from: srcName, fromId: source.id, port: source.name + ' → ' + dest.name });
          } catch { /* skip */ }
        }
      }
    }

    const result = [];
    for (let c = 0; c < modulesContainer.childNodes.length; c++) {
      const mod = modulesContainer.childNodes[c];
      const id = mod.id;
      if (!id) continue;
      const name = mod.dataset.name || '';
      const params = [];
      for (const input of mod.querySelectorAll('input')) {
        let label = '';
        const prev = input.previousSibling;
        if (prev && prev.textContent) label = prev.textContent.trim();
        if (input.type === 'checkbox') {
          params.push({ label, value: input.checked ? 'true' : 'false', type: 'checkbox' });
        } else {
          params.push({ label, value: input.value, type: input.type });
        }
      }
      for (const select of mod.querySelectorAll('select')) {
        let label = '';
        const prev = select.previousSibling;
        if (prev && prev.textContent) label = prev.textContent.trim();
        params.push({ label, value: select.value, type: 'select', options: Array.from(select.options).map(o => o.value) });
      }
      const buttons = [];
      for (const btn of mod.querySelectorAll('button')) {
        buttons.push(btn.textContent.trim());
      }
      const entry = { id, name, params, buttons };
      if (connections[id]) {
        entry.connectedFrom = connections[id].inputs;
        entry.connectedTo = connections[id].outputs;
      }
      result.push(entry);
    }
    return result;
  });
}

export async function getProgramLinks() {
  if (!page) throw new Error('Browser not launched');
  return page.evaluate(() => {
    const out = [];
    const svg = document.getElementById('svg');
    if (!svg) return out;
    const linksGroup = svg.getElementById('links');
    if (!linksGroup) return out;
    for (let l = 0; l < linksGroup.childNodes.length; l++) {
      const link = linksGroup.childNodes[l];
      if (!link.id) continue;
      try {
        const linkData = JSON.parse(link.id);
        const source = JSON.parse(linkData.source);
        const dest = JSON.parse(linkData.dest);
        const srcMod = document.getElementById(source.id);
        const destMod = document.getElementById(dest.id);
        const srcName = srcMod ? srcMod.dataset.name : source.id;
        const destName = destMod ? destMod.dataset.name : dest.id;
        out.push({ from: `${srcName}.${source.name}`, to: `${destName}.${dest.name}` });
      } catch { /* skip malformed link */ }
    }
    return out;
  });
}

export async function setModuleInput(moduleId, paramName, value) {
  if (!page) throw new Error('Browser not launched');
  return page.evaluate(({ moduleId, paramName, value }) => {
    const mod = document.getElementById(moduleId);
    if (!mod) return { error: `Module ${moduleId} not found` };
    const paramLower = paramName.toLowerCase();

    // Check <input> elements
    for (const input of mod.querySelectorAll('input')) {
      const prev = input.previousSibling;
      const label = prev ? prev.textContent.trim() : '';
      if (label.toLowerCase().includes(paramLower)) {
        if (input.type === 'checkbox') {
          input.checked = (value === 'true' || value === '1' || value === 'on');
          input.dispatchEvent(new Event('input', { bubbles: true }));
          input.dispatchEvent(new Event('change', { bubbles: true }));
          return { success: true, label, type: 'checkbox', newValue: input.checked };
        } else {
          input.value = value;
          input.dispatchEvent(new Event('input', { bubbles: true }));
          input.dispatchEvent(new Event('change', { bubbles: true }));
          return { success: true, label, newValue: value };
        }
      }
    }

    // Check <select> elements
    for (const select of mod.querySelectorAll('select')) {
      const prev = select.previousSibling;
      const label = prev ? prev.textContent.trim() : '';
      if (label.toLowerCase().includes(paramLower)) {
        select.value = value;
        select.dispatchEvent(new Event('input', { bubbles: true }));
        select.dispatchEvent(new Event('change', { bubbles: true }));
        return { success: true, label, type: 'select', newValue: select.value };
      }
    }

    return { error: `Parameter "${paramName}" not found in module ${moduleId}` };
  }, { moduleId, paramName, value: String(value) });
}

export async function clickModuleButton(moduleId, buttonText) {
  if (!page) throw new Error('Browser not launched');
  return page.evaluate(({ moduleId, buttonText }) => {
    const mod = document.getElementById(moduleId);
    if (!mod) return { error: `Module ${moduleId} not found` };
    for (const btn of mod.querySelectorAll('button')) {
      if (btn.textContent.trim().toLowerCase().includes(buttonText.toLowerCase())) {
        btn.click();
        return { success: true, clicked: btn.textContent.trim() };
      }
    }
    const available = Array.from(mod.querySelectorAll('button')).map(b => b.textContent.trim());
    return { error: `Button "${buttonText}" not found`, available };
  }, { moduleId, buttonText });
}

export async function injectProgram(programJson) {
  if (!page) throw new Error('Browser not launched');
  await page.evaluate((json) => {
    window.mods_prog_load(JSON.parse(json));
  }, JSON.stringify(programJson));
  await page.waitForFunction(() => {
    const modules = document.getElementById('modules');
    return modules && modules.childNodes.length > 0;
  }, { timeout: 15000 });
  await page.waitForTimeout(500);
}

export async function extractProgramState() {
  if (!page) throw new Error('Browser not launched');
  return page.evaluate(() => {
    if (typeof window.mods_build_v2_program === 'function') {
      return window.mods_build_v2_program();
    }
    // Fallback v1 extraction
    const prog = { modules: {}, links: [] };
    const modulesContainer = document.getElementById('modules');
    if (!modulesContainer) return null;
    for (let c = 0; c < modulesContainer.childNodes.length; c++) {
      const mod = modulesContainer.childNodes[c];
      if (!mod.id) continue;
      prog.modules[mod.id] = {
        definition: mod.dataset.definition || '',
        top: mod.dataset.top || '0',
        left: mod.dataset.left || '0',
        filename: mod.dataset.filename || '',
        inputs: {}, outputs: {}
      };
    }
    const svg = document.getElementById('svg');
    if (svg) {
      const links = svg.getElementById('links');
      if (links) {
        for (let l = 0; l < links.childNodes.length; l++) {
          if (links.childNodes[l].id) prog.links.push(links.childNodes[l].id);
        }
      }
    }
    return prog;
  });
}

export async function waitForModuleOutput({ outputName, moduleName, timeout = 10000 } = {}) {
  if (!page) throw new Error('Browser not launched');
  return page.evaluate(({ outputName, moduleName, timeout }) => {
    return new Promise(resolve => {
      const handler = (e) => {
        if (!e.data || e.data.type !== 'moduleOutput') return;
        if (outputName && e.data.output !== outputName) return;
        if (moduleName && e.data.module.name !== moduleName) return;
        window.removeEventListener('message', handler);
        resolve(e.data);
      };
      window.addEventListener('message', handler);
      setTimeout(() => { window.removeEventListener('message', handler); resolve(null); }, timeout);
    });
  }, { outputName, moduleName, timeout });
}

export async function waitForProcessingSignal({ timeout = 30000 } = {}) {
  if (!page) throw new Error('Browser not launched');
  // Race: moduleOutput event vs download vs timeout
  const beforeCount = downloads.length;
  const event = await page.evaluate((timeout) => {
    return new Promise(resolve => {
      const handler = (e) => {
        if (e.data && e.data.type === 'moduleOutput') {
          window.removeEventListener('message', handler);
          resolve(e.data);
        }
      };
      window.addEventListener('message', handler);
      setTimeout(() => { window.removeEventListener('message', handler); resolve(null); }, timeout);
    });
  }, Math.min(timeout, 3000)); // 3s max on old mods (no moduleOutput)
  // If no event, check if a download appeared during the wait
  if (!event && downloads.length > beforeCount) return { signal: 'download' };
  return event;
}

export async function getImageInfo() {
  if (!lastImageInfo || !lastImageInfo.width || !lastImageInfo.height) {
    return { error: 'No image loaded — use load_file first' };
  }
  return {
    moduleId: lastImageInfo.moduleId,
    moduleName: lastImageInfo.moduleName,
    pixelWidth: lastImageInfo.width,
    pixelHeight: lastImageInfo.height,
    currentDpi: lastImageInfo.dpi
  };
}

export function getLatestDownload() {
  return downloads.length > 0 ? downloads[downloads.length - 1] : null;
}

export function clearDownloads() {
  downloads = [];
}

export function getPage() {
  return page;
}

export function isLaunched() {
  return browserInstance !== null && page !== null;
}

export async function close() {
  if (browserInstance) {
    await browserInstance.close();
    browserInstance = null;
    page = null;
  }
}
