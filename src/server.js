#!/usr/bin/env node
// server.js — MCP server for remote mods CE interaction

import { stat, readFile, writeFile, mkdir, readdir } from 'node:fs/promises';
import { extname, join, dirname, basename } from 'node:path';
import { homedir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import vm from 'node:vm';
import * as browser from './browser.js';

// --- CLI ---
const args = process.argv.slice(2);

// Subcommands run before normal flag parsing so they can short-circuit.
if (args[0] === 'setup') {
  const r = spawnSync('npx', ['playwright', 'install', 'chromium'], { stdio: 'inherit' });
  process.exit(r.status ?? 1);
}
if (args[0] === '--help' || args[0] === '-h') {
  process.stdout.write([
    'Usage:',
    '  mops              Start the MCP server on stdio (default)',
    '  mops setup        Install the Playwright Chromium browser (run once)',
    '',
    'Flags:',
    '  --mods-url URL    Mods CE deployment to control (default: https://modsproject.org)',
    '  --headless        Run the browser without a visible window',
    '',
  ].join('\n'));
  process.exit(0);
}

let modsUrl = 'https://modsproject.org';
let headless = false;
for (let i = 0; i < args.length; i++) {
  if (args[i] === '--mods-url' && args[i + 1]) {
    modsUrl = args[i + 1];
    i++;
  }
  if (args[i] === '--headless') headless = true;
}
modsUrl = modsUrl.replace(/\/+$/, '');

// --- Manifest cache ---
let modulesManifest = null;
let programsManifest = null;

async function fetchManifest(type) {
  const url = `${modsUrl}/${type}/index.json`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Failed to fetch ${url}: ${res.status}`);
  return res.json();
}

async function getModulesManifest() {
  if (!modulesManifest) {
    const raw = await fetchManifest('modules');
    modulesManifest = raw.map(m => ({ ...m, path: decodeURIComponent(m.path) }));
  }
  return modulesManifest;
}

async function getProgramsManifest() {
  if (!programsManifest) {
    const raw = await fetchManifest('programs');
    programsManifest = raw.map(p => ({ ...p, path: decodeURIComponent(p.path) }));
  }
  return programsManifest;
}

function groupByCategory(items) {
  const groups = {};
  for (const item of items) {
    const cat = item.category || 'uncategorized';
    if (!groups[cat]) groups[cat] = [];
    groups[cat].push(item);
  }
  return groups;
}

// --- Module parsing (VM sandbox + regex fallback) ---

function extractWithVm(source) {
  const sandbox = {
    document: {
      createElement: () => ({
        style: {}, appendChild: () => {}, addEventListener: () => {},
        setAttribute: () => {}, getContext: () => ({
          canvas: { width: 0, height: 0 }, drawImage: () => {},
          getImageData: () => ({ data: [] }), putImageData: () => {},
          clearRect: () => {}, fillRect: () => {}, beginPath: () => {},
          moveTo: () => {}, lineTo: () => {}, stroke: () => {}, fill: () => {},
          arc: () => {}, closePath: () => {}, scale: () => {}, translate: () => {},
          save: () => {}, restore: () => {}, createImageData: () => ({ data: [] })
        }),
        classList: { add: () => {}, remove: () => {} },
        querySelectorAll: () => [], querySelector: () => null,
        removeChild: () => {}, insertBefore: () => {},
        children: [], childNodes: [], value: '', type: '', checked: false,
        innerHTML: '', textContent: '', createTextNode: () => ({})
      }),
      createTextNode: (t) => ({ textContent: t }),
      createElementNS: () => ({
        style: {}, appendChild: () => {}, setAttribute: () => {},
        addEventListener: () => {}, setAttributeNS: () => {},
        getBBox: () => ({ x: 0, y: 0, width: 0, height: 0 })
      }),
      getElementById: () => null,
      body: { appendChild: () => {}, removeChild: () => {} }
    },
    window: { addEventListener: () => {}, removeEventListener: () => {}, innerWidth: 800, innerHeight: 600 },
    mods: { ui: { padding: '5px', canvas: 200, header: 50, xstart: 0, ystart: 0 }, output: () => {}, input: () => {} },
    navigator: { userAgent: '', platform: '' },
    console: { log: () => {}, error: () => {} },
    SVGElement: function() {}, HTMLElement: function() {},
    Event: function() {}, CustomEvent: function() {},
    Blob: function() {}, URL: { createObjectURL: () => '' },
    FileReader: function() { this.readAsArrayBuffer = () => {}; this.readAsText = () => {}; this.onload = null; },
    WebSocket: function() { this.send = () => {}; this.close = () => {}; },
    XMLHttpRequest: function() { this.open = () => {}; this.send = () => {}; this.setRequestHeader = () => {}; },
    Image: function() {},
    Worker: function() { this.postMessage = () => {}; this.terminate = () => {}; },
    requestAnimationFrame: () => {}, setTimeout: () => {}, setInterval: () => {},
    clearTimeout: () => {}, clearInterval: () => {},
    parseInt, parseFloat, Math, JSON, Array, Object, String, Number, Boolean,
    RegExp, Date, Error, Map, Set, Promise, isNaN, isFinite,
    undefined, NaN, Infinity, encodeURIComponent, decodeURIComponent
  };
  const context = vm.createContext(sandbox);
  new vm.Script('var __result = ' + source).runInContext(context, { timeout: 1000 });
  return sandbox.__result;
}

function extractWithRegex(source) {
  const nameMatch = source.match(/var\s+name\s*=\s*['"]([^'"]+)['"]/);
  const name = nameMatch ? nameMatch[1] : 'unknown';
  const inputs = {};
  const inputsMatch = source.match(/var\s+inputs\s*=\s*\{([\s\S]*?)\n\s*\}/);
  if (inputsMatch) {
    for (const m of inputsMatch[1].matchAll(/(\w+)\s*:\s*\{[^}]*type\s*:\s*['"]([^'"]*)['"]/g)) {
      inputs[m[1]] = { type: m[2] };
    }
  }
  const outputs = {};
  const outputsMatch = source.match(/var\s+outputs\s*=\s*\{([\s\S]*?)\n\s*\}/);
  if (outputsMatch) {
    for (const m of outputsMatch[1].matchAll(/(\w+)\s*:\s*\{[^}]*type\s*:\s*['"]([^'"]*)['"]/g)) {
      outputs[m[1]] = { type: m[2] };
    }
  }
  return { name, inputs, outputs };
}

async function parseModule(modulePath, includeSource) {
  const url = `${modsUrl}/${modulePath}`;
  const res = await fetch(url);
  if (!res.ok) return { path: modulePath, error: `HTTP ${res.status} fetching ${url}`, parseMethod: 'failed' };
  const source = await res.text();

  let name, inputs, outputs, parseMethod = 'vm';
  try {
    const result = extractWithVm(source);
    name = result.name;
    inputs = {};
    outputs = {};
    if (result.inputs) for (const [k, v] of Object.entries(result.inputs)) inputs[k] = { type: v.type || '' };
    if (result.outputs) for (const [k, v] of Object.entries(result.outputs)) outputs[k] = { type: v.type || '' };
  } catch {
    parseMethod = 'regex';
    try {
      const result = extractWithRegex(source);
      name = result.name;
      inputs = result.inputs;
      outputs = result.outputs;
    } catch (err) {
      return { path: modulePath, error: `Failed to parse: ${err.message}`, parseMethod: 'failed' };
    }
  }

  const info = { name, path: modulePath, inputs, outputs, parseMethod };
  if (includeSource) info.source = source;
  return info;
}

// --- State ---
let loadedProgram = null;
let lastLoadedFile = null;

// --- MCP Server ---
const mcpServer = new McpServer({ name: 'mops', version: '0.4.0' });

// Wrap tool registration so every tool call is logged with args, duration,
// and result/error. Output goes to stderr (MCP convention) and is captured
// by mops-voice into its session log file.
const _origTool = mcpServer.tool.bind(mcpServer);
mcpServer.tool = function (name, description, schema, handler) {
  const wrapped = async (args) => {
    const ts = new Date().toISOString().slice(11, 23);
    const argStr = (() => {
      try { return JSON.stringify(args ?? {}).slice(0, 300); }
      catch { return '<unserializable>'; }
    })();
    console.error(`[mops][tool] ${ts} → ${name} ${argStr}`);
    const t0 = Date.now();
    try {
      const result = await handler(args);
      const ms = Date.now() - t0;
      const text = result?.content?.[0]?.text ?? '';
      const preview = String(text).slice(0, 240).replace(/\s+/g, ' ');
      const tag = result?.isError ? 'ERR' : 'OK';
      console.error(`[mops][tool] ← ${name} [${tag}] ${ms}ms: ${preview}`);
      return result;
    } catch (err) {
      const ms = Date.now() - t0;
      console.error(`[mops][tool] ✖ ${name} ${ms}ms: ${err.stack || err.message}`);
      throw err;
    }
  };
  return _origTool(name, description, schema, wrapped);
};

async function findModule(moduleName, moduleId) {
  const state = await browser.getProgramState();
  if (moduleId) {
    const mod = state.find(m => m.id === moduleId);
    return mod ? { module: mod } : { error: `Module with ID "${moduleId}" not found.` };
  }
  const mod = state.find(m => m.name.toLowerCase().includes(moduleName.toLowerCase()));
  if (!mod) {
    const available = state.map(m => m.name).filter(Boolean);
    return { error: `Module "${moduleName}" not found. Available: ${available.join(', ')}` };
  }
  return { module: mod };
}

async function buildProgramSnapshot() {
  const [state, links] = await Promise.all([browser.getProgramState(), browser.getProgramLinks()]);
  return {
    modules: state.map(m => ({
      id: m.id,
      name: m.name,
      params: m.params.map(p => ({ label: p.label, value: p.value })),
      buttons: m.buttons,
    })),
    links,
  };
}

function parseModuleNameId(module_name) {
  let name = module_name, id = undefined;
  if (module_name.includes(':0.')) {
    const parts = module_name.split(':');
    name = parts[0];
    id = parts.slice(1).join(':');
  }
  return { name, id };
}

// --- User profile ---
const PROFILE_DIR = join(homedir(), '.mops');
const PROFILE_PATH = join(PROFILE_DIR, 'profile.json');

async function loadProfile() {
  try {
    const data = await readFile(PROFILE_PATH, 'utf-8');
    return JSON.parse(data);
  } catch {
    return { machines: [], preferences: {} };
  }
}

async function saveProfile(profile) {
  await mkdir(PROFILE_DIR, { recursive: true });
  await writeFile(PROFILE_PATH, JSON.stringify(profile, null, 2));
}

// --- Tools ---

mcpServer.tool('get_server_status', 'Get server health, browser state, mods URL, and loaded program', {},
  async () => {
    const status = {
      server: 'running', modsUrl,
      browser: browser.isLaunched() ? 'connected' : 'not launched',
      loadedProgram: loadedProgram || 'none'
    };
    if (browser.isLaunched() && loadedProgram) {
      try {
        const state = await browser.getProgramState();
        status.moduleCount = state.length;
        status.moduleNames = state.map(m => m.name).filter(Boolean);
      } catch { /* ignore */ }
    }
    return { content: [{ type: 'text', text: JSON.stringify(status, null, 2) }] };
  }
);

mcpServer.tool('get_profile', 'Get user profile: machines, preferences, and other saved settings', {},
  async () => {
    const profile = await loadProfile();
    if (profile.machines.length === 0 && Object.keys(profile.preferences).length === 0) {
      return { content: [{ type: 'text', text: 'No profile configured yet. Use update_profile to add your machines and preferences.' }] };
    }
    return { content: [{ type: 'text', text: JSON.stringify(profile, null, 2) }] };
  }
);

mcpServer.tool('update_profile',
  'Add, remove, or update machines and preferences in the user profile. Stored locally at ~/.mops/profile.json.',
  {
    action: z.enum(['add_machine', 'remove_machine', 'set_preference', 'remove_preference']).describe('What to do'),
    machine: z.object({
      name: z.string().describe('Machine name (e.g., "Roland GX-24")'),
      type: z.string().describe('What it does (e.g., "vinyl cutter", "CNC mill", "3D printer", "laser cutter")'),
      program: z.string().optional().describe('Mods program path if known (e.g., "programs/machines/Roland/GX-24/cut vinyl")'),
      deviceName: z.string().optional().describe('USB/serial device name as shown in Chrome device picker (e.g., "Roland DG SRM-20"). Used for WebUSB/WebSerial auto-selection.'),
      notes: z.string().optional().describe('Any extra info (e.g., "max area 24x12 inches", "in room 302")')
    }).optional().describe('Machine details (for add_machine/remove_machine)'),
    preference: z.object({
      key: z.string().describe('Preference name (e.g., "default_units", "output_directory")'),
      value: z.string().describe('Preference value')
    }).optional().describe('Preference key-value (for set_preference/remove_preference)')
  },
  async ({ action, machine, preference }) => {
    const profile = await loadProfile();

    if (action === 'add_machine') {
      if (!machine) return { content: [{ type: 'text', text: 'Error: machine is required for add_machine' }], isError: true };
      const existing = profile.machines.findIndex(m => m.name.toLowerCase() === machine.name.toLowerCase());
      if (existing >= 0) profile.machines[existing] = machine;
      else profile.machines.push(machine);
      await saveProfile(profile);
      return { content: [{ type: 'text', text: `Machine "${machine.name}" saved. You now have ${profile.machines.length} machine(s).` }] };
    }

    if (action === 'remove_machine') {
      if (!machine) return { content: [{ type: 'text', text: 'Error: machine is required for remove_machine' }], isError: true };
      const before = profile.machines.length;
      profile.machines = profile.machines.filter(m => m.name.toLowerCase() !== machine.name.toLowerCase());
      if (profile.machines.length === before) return { content: [{ type: 'text', text: `Machine "${machine.name}" not found in profile.` }] };
      await saveProfile(profile);
      return { content: [{ type: 'text', text: `Machine "${machine.name}" removed. ${profile.machines.length} machine(s) remaining.` }] };
    }

    if (action === 'set_preference') {
      if (!preference) return { content: [{ type: 'text', text: 'Error: preference is required for set_preference' }], isError: true };
      profile.preferences[preference.key] = preference.value;
      await saveProfile(profile);
      return { content: [{ type: 'text', text: `Preference "${preference.key}" set to "${preference.value}".` }] };
    }

    if (action === 'remove_preference') {
      if (!preference) return { content: [{ type: 'text', text: 'Error: preference is required for remove_preference' }], isError: true };
      if (!(preference.key in profile.preferences)) return { content: [{ type: 'text', text: `Preference "${preference.key}" not found.` }] };
      delete profile.preferences[preference.key];
      await saveProfile(profile);
      return { content: [{ type: 'text', text: `Preference "${preference.key}" removed.` }] };
    }
  }
);

mcpServer.tool('find_machine',
  'Find the best matching machine from the user profile for a given task, and match it to an available Mods program.',
  {
    task: z.string().describe('What the user wants to do (e.g., "cut a sticker", "mill a PCB", "3D print a case")')
  },
  async ({ task }) => {
    const profile = await loadProfile();
    if (profile.machines.length === 0) {
      return { content: [{ type: 'text', text: 'No machines in profile. Use update_profile to add your machines first.' }] };
    }
    const programs = await getProgramsManifest();
    const taskLower = task.toLowerCase();

    // Build machine results with matching programs
    const results = [];
    for (const machine of profile.machines) {
      const machineResult = { ...machine, matchingPrograms: [] };

      // If machine has a program path set, include it directly
      if (machine.program) {
        const found = programs.find(p => p.path === machine.program);
        if (found) machineResult.matchingPrograms.push(found);
      }

      // Search programs by machine name keywords.
      // Split on spaces only — model numbers like "GX-24" and "SRM-20" must stay
      // as single tokens so they match model-specific program paths.
      const keywords = machine.name.toLowerCase().split(/\s+/).filter(w => w.length > 2);
      for (const prog of programs) {
        const progPath = prog.path.toLowerCase();
        const progName = (prog.name || '').toLowerCase();
        const match = keywords.some(kw => progPath.includes(kw) || progName.includes(kw));
        if (match && !machineResult.matchingPrograms.some(p => p.path === prog.path)) {
          machineResult.matchingPrograms.push(prog);
        }
      }
      results.push(machineResult);
    }

    // Score relevance to the task
    const taskKeywords = {
      'sticker': ['vinyl', 'cut vinyl', 'cutter'],
      'vinyl': ['vinyl', 'cut vinyl', 'cutter'],
      'cut': ['cut', 'laser', 'vinyl', 'cutter'],
      'mill': ['mill', 'pcb', 'milling', 'cnc'],
      'pcb': ['mill', 'pcb', 'traces', 'outline'],
      'print': ['print', '3d', 'printer'],
      '3d': ['print', '3d', 'printer'],
      'engrave': ['laser', 'engrave'],
      'laser': ['laser', 'cut', 'engrave'],
      'scan': ['scan', 'scanner'],
      'route': ['route', 'router', 'cnc', 'shopbot']
    };

    // Find machines whose type matches the task
    const scored = results.map(m => {
      let score = 0;
      const mType = m.type.toLowerCase();
      const mName = m.name.toLowerCase();
      const mNotes = (m.notes || '').toLowerCase();
      for (const word of taskLower.split(/\s+/)) {
        if (mType.includes(word) || mName.includes(word) || mNotes.includes(word)) score += 2;
        const related = taskKeywords[word] || [];
        for (const r of related) {
          if (mType.includes(r) || mName.includes(r) || mNotes.includes(r)) score += 1;
        }
      }
      return { ...m, relevanceScore: score };
    });

    scored.sort((a, b) => b.relevanceScore - a.relevanceScore);

    return { content: [{ type: 'text', text: JSON.stringify(scored, null, 2) }] };
  }
);

mcpServer.tool('launch_browser',
  'Launch the Chromium browser and navigate to the mods CE deployment. Must be called before browser-dependent tools. Automatically sets up WebUSB/WebSerial device auto-selection from the user profile.',
  {},
  async () => {
    if (browser.isLaunched()) return { content: [{ type: 'text', text: `Browser already running at ${modsUrl}` }] };
    try {
      // Load device name filters from profile before launching
      const profile = await loadProfile();
      const filters = profile.machines
        .filter(m => m.deviceName)
        .map(m => m.deviceName);
      const names = profile.machines.map(m => m.name);
      browser.setDeviceFilters(filters, names);

      await browser.launch(modsUrl, headless);
      const msg = `Browser launched (${headless ? 'headless' : 'headed'}) at ${modsUrl}`;
      const deviceMsg = `. Device auto-select enabled for ${names.length} machine(s): ${names.join(', ')}`;
      return { content: [{ type: 'text', text: msg + deviceMsg }] };
    } catch (err) {
      return { content: [{ type: 'text', text: `Browser launch failed: ${err.message}. Run "npx playwright install chromium" to install browsers.` }], isError: true };
    }
  }
);

mcpServer.tool('list_devices',
  'List connected USB/serial devices. Shows devices discovered from picker prompts and devices granted via WebUSB. Use after a workflow to see what the user connected.',
  {},
  async () => {
    if (!browser.isLaunched()) return { content: [{ type: 'text', text: 'Error: Browser not launched.' }], isError: true };

    // Get devices from CDP prompts
    const discovered = browser.getDiscoveredDevices();

    // Also query granted WebUSB devices from the browser
    const granted = await browser.getGrantedDevices();

    const result = {};
    if (discovered.length > 0) result.discoveredInPrompts = discovered;
    if (granted.length > 0) result.grantedUsbDevices = granted;

    if (discovered.length === 0 && granted.length === 0) {
      return { content: [{ type: 'text', text: 'No devices found. Devices appear after a workflow triggers WebUSB/WebSerial (e.g., sending output to a machine).' }] };
    }
    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  }
);

mcpServer.tool('list_programs', 'List available Mods programs organized by category',
  { category: z.string().optional().describe('Filter by category (e.g., "machines", "processes", "image")') },
  async ({ category }) => {
    try {
      const manifest = await getProgramsManifest();
      let items = manifest;
      if (category) items = manifest.filter(p => p.category.toLowerCase().includes(category.toLowerCase()));
      return { content: [{ type: 'text', text: JSON.stringify(groupByCategory(items), null, 2) }] };
    } catch (err) {
      return { content: [{ type: 'text', text: `Error fetching programs: ${err.message}` }], isError: true };
    }
  }
);

mcpServer.tool('list_modules', 'List available Mods modules organized by category',
  { category: z.string().optional().describe('Filter by category (e.g., "read", "image", "mesh")') },
  async ({ category }) => {
    try {
      const manifest = await getModulesManifest();
      let items = manifest;
      if (category) items = manifest.filter(m => m.category.toLowerCase().includes(category.toLowerCase()));
      return { content: [{ type: 'text', text: JSON.stringify(groupByCategory(items), null, 2) }] };
    } catch (err) {
      return { content: [{ type: 'text', text: `Error fetching modules: ${err.message}` }], isError: true };
    }
  }
);

mcpServer.tool('get_module_info',
  'Parse module file(s) and return name, inputs, outputs with types. Accepts a single path or array.',
  {
    path: z.union([z.string(), z.array(z.string())]).describe('Module path(s) (e.g., "modules/read/stl.js" or ["modules/read/svg.js", "modules/read/png.js"])'),
    include_source: z.boolean().optional().default(false).describe('Include full IIFE source in response')
  },
  async ({ path, include_source }) => {
    const paths = Array.isArray(path) ? path : [path];
    const results = await Promise.all(paths.map(p => parseModule(p, include_source)));
    return { content: [{ type: 'text', text: JSON.stringify(results.length === 1 ? results[0] : results, null, 2) }] };
  }
);

mcpServer.tool('load_program',
  'Load a preset program in the browser by path. Optionally preload a file into the matching reader module via src URL.',
  {
    path: z.string().describe('Program path (e.g., "programs/machines/Roland/SRM-20 mill/mill 2D PCB")'),
    src: z.string().optional().describe('URL of a file to auto-load into the matching reader module (matched by extension)')
  },
  async ({ path, src }) => {
    if (!browser.isLaunched()) return { content: [{ type: 'text', text: 'Error: Browser not launched. Use launch_browser first.' }], isError: true };
    await browser.loadProgram(modsUrl, path, src);
    loadedProgram = path;
    // Navigating to a new program invalidates any prior file state
    lastLoadedFile = src || null;
    const snapshot = await buildProgramSnapshot();
    const result = { loaded: path, ...snapshot };
    if (src) result.src = src;
    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  }
);

mcpServer.tool('get_program_state', 'Get current state of all modules in the loaded program', {},
  async () => {
    if (!browser.isLaunched()) return { content: [{ type: 'text', text: 'Error: Browser not launched.' }], isError: true };
    if (!loadedProgram) return { content: [{ type: 'text', text: 'Error: No program loaded. Use load_program first.' }], isError: true };
    const state = await browser.getProgramState();
    return { content: [{ type: 'text', text: JSON.stringify(state, null, 2) }] };
  }
);

mcpServer.tool('set_parameter', 'Set a parameter value in a specific module',
  {
    module_name: z.string().describe('Module name (or partial match, or name:id for disambiguation)'),
    parameter: z.string().describe('Parameter label (or partial match)'),
    value: z.string().describe('New value to set')
  },
  async ({ module_name, parameter, value }) => {
    if (!browser.isLaunched()) return { content: [{ type: 'text', text: 'Error: Browser not launched.' }], isError: true };
    const { name, id } = parseModuleNameId(module_name);
    const found = await findModule(name, id);
    if (found.error) return { content: [{ type: 'text', text: found.error }], isError: true };
    const result = await browser.setModuleInput(found.module.id, parameter, value);
    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  }
);

mcpServer.tool('set_parameters', 'Set multiple parameters across one or more modules in a single call. More efficient than multiple set_parameter calls.',
  {
    params: z.array(z.object({
      module_name: z.string().describe('Module name (or partial match, or name:id for disambiguation)'),
      parameter: z.string().describe('Parameter label (or partial match)'),
      value: z.string().describe('New value to set')
    })).describe('Array of parameters to set')
  },
  async ({ params }) => {
    if (!browser.isLaunched()) return { content: [{ type: 'text', text: 'Error: Browser not launched.' }], isError: true };
    const results = [];
    for (const { module_name, parameter, value } of params) {
      const { name, id } = parseModuleNameId(module_name);
      const found = await findModule(name, id);
      if (found.error) { results.push({ module_name, parameter, error: found.error }); continue; }
      const result = await browser.setModuleInput(found.module.id, parameter, value);
      results.push({ module_name, parameter, ...result });
    }
    return { content: [{ type: 'text', text: JSON.stringify(results, null, 2) }] };
  }
);

mcpServer.tool('trigger_action', 'Click a button in a module (calculate, view, export, etc.)',
  {
    module_name: z.string().describe('Module name (or partial match, or name:id for disambiguation)'),
    action: z.string().describe('Button text to click (or partial match)')
  },
  async ({ module_name, action }) => {
    if (!browser.isLaunched()) return { content: [{ type: 'text', text: 'Error: Browser not launched.' }], isError: true };
    browser.clearDownloads();
    const { name, id } = parseModuleNameId(module_name);
    const found = await findModule(name, id);
    if (found.error) return { content: [{ type: 'text', text: found.error }], isError: true };
    const result = await browser.clickModuleButton(found.module.id, action);
    // Wait for processing signal (moduleOutput on new mods, download or 3s fallback on old)
    const outputEvent = await browser.waitForProcessingSignal({ timeout: 30000 });
    if (outputEvent) {
      result.completedModule = outputEvent.module;
      result.completedOutput = outputEvent.output;
    }
    const download = browser.getLatestDownload();
    if (download) result.download = { filename: download.suggestedFilename, size: download.content ? download.content.length : 0 };
    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  }
);

mcpServer.tool('load_file',
  'Load a file into the matching reader module. Uses postMessage for SVG/PNG, file input for other types.',
  {
    module_name: z.string().describe('Module name (or partial match, or name:id for disambiguation)'),
    file_path: z.string().describe('Absolute path to the file to load')
  },
  async ({ module_name, file_path }) => {
    if (!browser.isLaunched()) return { content: [{ type: 'text', text: 'Error: Browser not launched.' }], isError: true };
    try { await stat(file_path); } catch {
      const ctx = await buildMissingFileContext(file_path);
      return {
        content: [{ type: 'text', text: JSON.stringify({ error: `File not found: ${file_path}`, ...ctx }, null, 2) }],
        isError: true
      };
    }
    lastLoadedFile = file_path;
    const ext = extname(file_path).toLowerCase();
    if (ext === '.svg' || ext === '.png') {
      const result = await browser.postMessageFile(file_path);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    const { name, id } = parseModuleNameId(module_name);
    const found = await findModule(name, id);
    if (found.error) return { content: [{ type: 'text', text: found.error }], isError: true };
    const result = await browser.setModuleFile(found.module.id, file_path);
    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  }
);

async function readImageDimensions(filePath) {
  const ext = extname(filePath).toLowerCase();
  const buf = await readFile(filePath);
  if (ext === '.png') {
    // PNG IHDR chunk: width at byte 16, height at byte 20 (4-byte big-endian)
    return { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) };
  }
  // SVG: parse width/height or viewBox
  if (ext === '.svg') {
    const svg = buf.toString('utf-8');
    const vbMatch = svg.match(/viewBox=["'][\d.]+\s+[\d.]+\s+([\d.]+)\s+([\d.]+)["']/);
    if (vbMatch) return { width: parseFloat(vbMatch[1]), height: parseFloat(vbMatch[2]) };
    const wMatch = svg.match(/width=["']([\d.]+)/);
    const hMatch = svg.match(/height=["']([\d.]+)/);
    if (wMatch && hMatch) return { width: parseFloat(wMatch[1]), height: parseFloat(hMatch[1]) };
  }
  return null;
}

mcpServer.tool('set_physical_size',
  'Set the physical output size for a loaded PNG image. Reads pixel dimensions from the file header, calculates the correct DPI, and sets it on the reader module. Only works with PNG — vector formats (DXF, HPGL, SVG) have physical dimensions defined by their data.',
  {
    width: z.number().describe('Desired physical width'),
    height: z.number().optional().describe('Desired physical height (if omitted, aspect ratio is preserved from width)'),
    unit: z.enum(['mm', 'cm', 'in']).default('mm').describe('Unit for width/height')
  },
  async ({ width, height, unit }) => {
    if (!browser.isLaunched()) return { content: [{ type: 'text', text: 'Error: Browser not launched.' }], isError: true };
    if (!loadedProgram) return { content: [{ type: 'text', text: 'Error: No program loaded.' }], isError: true };
    if (!lastLoadedFile) return { content: [{ type: 'text', text: 'Error: No image file loaded. Use load_file first.' }], isError: true };

    // Read pixel dimensions directly from the file (PNG only — DPI controls physical size)
    const ext = extname(lastLoadedFile).toLowerCase();
    if (ext !== '.png') {
      return { content: [{ type: 'text', text: `Error: set_physical_size only works with PNG files. For vector formats (DXF, HPGL, SVG), physical dimensions come from the file data — DPI only controls rasterization resolution.` }], isError: true };
    }
    const dims = await readImageDimensions(lastLoadedFile);
    if (!dims) return { content: [{ type: 'text', text: `Error: Cannot read dimensions from ${lastLoadedFile}` }], isError: true };

    // Find the reader module — try name first, fall back to any module with DPI
    const readerNames = ['read', 'png', 'svg', 'image'];
    let found = null;
    for (const name of readerNames) {
      found = await findModule(name);
      if (!found.error && found.module.params.some(p => p.label.toLowerCase().includes('dpi'))) break;
      found = null;
    }
    const dpiModule = found ? found.module : null;
    if (!dpiModule) return { content: [{ type: 'text', text: 'Error: No reader module with a DPI parameter found in the loaded program.' }], isError: true };

    // Calculate exact DPI from desired width
    const widthInches = unit === 'mm' ? width / 25.4 : unit === 'cm' ? width / 2.54 : width;
    const newDpi = dims.width / widthInches;

    // Set DPI (3 decimal places avoids floating point noise)
    const setResult = await browser.setModuleInput(dpiModule.id, 'dpi', newDpi.toFixed(3));
    if (setResult.error) return { content: [{ type: 'text', text: `Error setting DPI: ${setResult.error}` }], isError: true };

    // Confirmation with resulting dimensions
    const rW_mm = (25.4 * dims.width / newDpi).toFixed(1);
    const rH_mm = (25.4 * dims.height / newDpi).toFixed(1);
    const rW_in = (dims.width / newDpi).toFixed(3);
    const rH_in = (dims.height / newDpi).toFixed(3);

    const response = {
      success: true, dpiSet: parseFloat(newDpi.toFixed(3)),
      imagePixels: `${dims.width} x ${dims.height}`,
      resultingSize: { mm: `${rW_mm} x ${rH_mm} mm`, in: `${rW_in} x ${rH_in} in` }
    };

    if (height !== undefined) {
      const heightInches = unit === 'mm' ? height / 25.4 : unit === 'cm' ? height / 2.54 : height;
      const actualHeightInches = dims.height / newDpi;
      if (Math.abs(actualHeightInches - heightInches) / heightInches > 0.05) {
        const actualH = actualHeightInches * (unit === 'mm' ? 25.4 : unit === 'cm' ? 2.54 : 1);
        response.warning = `Image aspect ratio doesn't match. Height will be ${actualH.toFixed(1)} ${unit} instead of ${height} ${unit}. DPI is set based on width.`;
      }
    }

    return { content: [{ type: 'text', text: JSON.stringify(response, null, 2) }] };
  }
);

mcpServer.tool('export_file', 'Get the most recently downloaded/exported file from Mods', {},
  async () => {
    const download = browser.getLatestDownload();
    if (!download) return { content: [{ type: 'text', text: 'No file exported yet. Use trigger_action first.' }], isError: true };
    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          filename: download.suggestedFilename,
          size: download.content ? download.content.length : 0,
          content: download.content ? download.content.toString('utf-8').slice(0, 10000) : null
        }, null, 2)
      }]
    };
  }
);

mcpServer.tool('create_program', 'Build a new v2 program from modules and connections, load in browser',
  {
    modules: z.array(z.string()).describe('Module paths (e.g., ["modules/read/svg.js", "modules/mesh/rotate.js"])'),
    links: z.array(z.object({
      from: z.string().describe('Source: "moduleName.outputPort"'),
      to: z.string().describe('Destination: "moduleName.inputPort"')
    })).describe('Connections between modules')
  },
  async ({ modules: modulePaths, links }) => {
    if (!browser.isLaunched()) return { content: [{ type: 'text', text: 'Error: Browser not launched.' }], isError: true };
    try {
      const mods = {};
      const nameToId = {};
      let x = 100;
      for (const modPath of modulePaths) {
        const id = Math.random().toString();
        mods[id] = { module: modPath, top: '100', left: String(x), params: {} };
        x += 300;
        const res = await fetch(`${modsUrl}/${modPath}`);
        if (res.ok) {
          const source = await res.text();
          const nameMatch = source.match(/var\s+name\s*=\s*['"]([^'"]+)['"]/);
          if (nameMatch) nameToId[nameMatch[1]] = id;
        }
      }
      const programLinks = [];
      for (const link of links) {
        const [fromModule, fromPort] = link.from.split('.');
        const [toModule, toPort] = link.to.split('.');
        const sourceId = nameToId[fromModule];
        const destId = nameToId[toModule];
        if (!sourceId || !destId) throw new Error(`Module not found in link: ${!sourceId ? fromModule : toModule}`);
        programLinks.push(JSON.stringify({
          source: JSON.stringify({ id: sourceId, type: 'outputs', name: fromPort }),
          dest: JSON.stringify({ id: destId, type: 'inputs', name: toPort })
        }));
      }
      const prog = { version: 2, modules: mods, links: programLinks };
      await browser.injectProgram(prog);
      loadedProgram = 'custom (created)';
      const state = await browser.getProgramState();
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            created: true, moduleCount: Object.keys(mods).length,
            linkCount: programLinks.length, modules: state.map(m => ({ id: m.id, name: m.name }))
          }, null, 2)
        }]
      };
    } catch (err) {
      return { content: [{ type: 'text', text: `Error creating program: ${err.message}` }], isError: true };
    }
  }
);

mcpServer.tool('save_program', 'Extract the current program state as v2 JSON', {},
  async () => {
    if (!browser.isLaunched()) return { content: [{ type: 'text', text: 'Error: Browser not launched.' }], isError: true };
    const programState = await browser.extractProgramState();
    if (!programState) return { content: [{ type: 'text', text: 'Error: Could not extract program state.' }], isError: true };
    return { content: [{ type: 'text', text: JSON.stringify(programState, null, 2) }] };
  }
);

// --- Composite tools: shared helpers ---

const MACHINE_PARAMS_PATH = fileURLToPath(new URL('./machine-params.json', import.meta.url));

async function loadMachineParams() {
  try {
    const raw = await readFile(MACHINE_PARAMS_PATH, 'utf-8');
    return JSON.parse(raw);
  } catch {
    return { machines: {} };
  }
}

function upstreamOf(state, moduleId) {
  const mod = state.find(m => m.id === moduleId);
  if (!mod?.connectedFrom) return [];
  return mod.connectedFrom
    .map(link => state.find(m => m.id === link.fromId))
    .filter(Boolean);
}

// The module that talks to the machine. Identified by a button label that
// only ever appears on WebUSB/WebSerial output modules ("get device", "send file",
// "waiting for file"). Name-based matching is unreliable across program variants.
function findOutputSendModule(state) {
  const sendLabels = /send file|get device|waiting for file|ready to send/i;
  return state.find(m => m.buttons?.some(b => sendLabels.test(b))) || null;
}

function findCalculateModule(state) {
  return state.find(m => m.buttons?.some(b => /^\s*calculate\s*$/i.test(b))) || null;
}

// Per Fran: an on/off gate, when present, sits immediately upstream of the
// module it gates. Look one hop only — no recursive walk.
function findOnOffGate(state, targetModuleId) {
  return upstreamOf(state, targetModuleId).find(m => /on\/off/i.test(m.name)) || null;
}

function findReaderModule(state) {
  const readerNames = ['read', 'png', 'svg', 'image'];
  for (const kw of readerNames) {
    const m = state.find(
      mod => mod.name.toLowerCase().includes(kw)
        && mod.params.some(p => p.label.toLowerCase().includes('dpi'))
    );
    if (m) return m;
  }
  return state.find(m => m.params.some(p => p.label.toLowerCase().includes('dpi'))) || null;
}

function matchProfileMachine(profile, hint) {
  if (!hint || !profile.machines?.length) return null;
  const hintLower = hint.toLowerCase();
  // Exact/substring match on name
  let m = profile.machines.find(
    x => x.name.toLowerCase() === hintLower || x.name.toLowerCase().includes(hintLower)
  );
  if (m) return m;
  // Match on type (e.g., "vinyl cutter")
  m = profile.machines.find(x => (x.type || '').toLowerCase().includes(hintLower));
  if (m) return m;
  // Keyword match against name+type+notes
  m = profile.machines.find(x => {
    const bag = `${x.name} ${x.type || ''} ${x.notes || ''}`.toLowerCase();
    return hintLower.split(/\s+/).some(w => w.length > 2 && bag.includes(w));
  });
  return m || null;
}

function composeError(at_step, reason, extra = {}) {
  return {
    content: [{ type: 'text', text: JSON.stringify({ status: 'failed', at_step, reason, ...extra }, null, 2) }],
    isError: true
  };
}

function composeOk(payload) {
  return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] };
}

// On file-not-found, surface the parent directory contents so the LLM can
// recover from near-miss filenames (e.g. "version 2.png" vs "version2.png")
// instead of the server guessing with ad-hoc heuristics.
async function buildMissingFileContext(file_path) {
  const dir = dirname(file_path);
  const wanted = basename(file_path);
  let entries = [];
  let directory_exists = true;
  try {
    entries = (await readdir(dir)).filter(f => !f.startsWith('.'));
  } catch {
    directory_exists = false;
  }
  const MAX_ENTRIES = 50;
  const truncated = entries.length > MAX_ENTRIES;
  return {
    wanted_basename: wanted,
    directory: dir,
    directory_exists,
    entries: entries.slice(0, MAX_ENTRIES),
    entries_truncated: truncated,
    total_entries: entries.length,
    hint: directory_exists
      ? 'If one of these is clearly the intended file, retry with its exact name. If ambiguous, ask the user which one they meant. If none match, tell the user the file isn\'t there.'
      : 'The directory itself does not exist. Ask the user where the file actually lives.'
  };
}

// Shared DPI-from-physical-size logic used by both set_physical_size and setup_cut.
async function applyPhysicalSize(state, filePath, size) {
  const ext = extname(filePath).toLowerCase();
  if (ext !== '.png') {
    return { error: 'set_physical_size only works with PNG files' };
  }
  const buf = await readFile(filePath);
  const dims = { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) };
  const reader = findReaderModule(state);
  if (!reader) return { error: 'No reader module with a DPI parameter found' };
  const { width, height, unit } = size;
  const widthInches = unit === 'mm' ? width / 25.4 : unit === 'cm' ? width / 2.54 : width;
  const newDpi = dims.width / widthInches;
  const r = await browser.setModuleInput(reader.id, 'dpi', newDpi.toFixed(3));
  if (r.error) return { error: r.error };
  const result = {
    dpi: parseFloat(newDpi.toFixed(3)),
    pixels: `${dims.width} x ${dims.height}`,
    mm: `${(25.4 * dims.width / newDpi).toFixed(1)} x ${(25.4 * dims.height / newDpi).toFixed(1)}`
  };
  if (height !== undefined) {
    const heightInches = unit === 'mm' ? height / 25.4 : unit === 'cm' ? height / 2.54 : height;
    const actualHeightInches = dims.height / newDpi;
    if (Math.abs(actualHeightInches - heightInches) / heightInches > 0.05) {
      const actualH = actualHeightInches * (unit === 'mm' ? 25.4 : unit === 'cm' ? 2.54 : 1);
      result.warning = `Aspect mismatch: height will be ${actualH.toFixed(1)} ${unit}, not ${height} ${unit}`;
    }
  }
  return result;
}

// --- get_job_status ---

mcpServer.tool('get_job_status',
  'Read-only snapshot of current session: machine, program, file, size, device, toolpath readiness, next step. Call at the start of a turn to skip re-probing state.',
  {},
  async () => {
    const snapshot = {
      browser_connected: browser.isLaunched(),
      program_loaded: loadedProgram || null,
      file_loaded: lastLoadedFile || null,
      device_connected: false,
      device_name: null,
      toolpath_calculated: false,
      ready_to_send: false,
      next_step: null
    };

    if (!browser.isLaunched()) {
      snapshot.next_step = 'launch_browser';
      return composeOk(snapshot);
    }

    const granted = await browser.getGrantedDevices();
    if (granted.length > 0) {
      snapshot.device_connected = true;
      snapshot.device_name = granted[0].name;
    }

    if (!loadedProgram) {
      snapshot.next_step = 'load_program';
      return composeOk(snapshot);
    }

    const state = await browser.getProgramState();

    // Machine: fuzzy-match profile against program path
    const profile = await loadProfile();
    const pathLower = loadedProgram.toLowerCase();
    const machine = profile.machines?.find(m =>
      m.name.toLowerCase().split(/\s+/).filter(w => w.length > 2)
        .some(kw => pathLower.includes(kw))
    );
    if (machine) snapshot.machine = machine.name;

    // Non-default-looking params, flat — label collisions tolerated (rare in practice)
    const params = {};
    for (const mod of state) {
      for (const p of mod.params) {
        if (p.value === '' || p.value === null || p.value === 'false') continue;
        params[p.label] = p.value;
      }
    }
    snapshot.parameters = params;

    // Output + toolpath
    const outputMod = findOutputSendModule(state);
    if (outputMod) {
      const sendLabels = /send file|get device|waiting for file|ready to send/i;
      const label = outputMod.buttons.find(b => sendLabels.test(b)) || outputMod.buttons[0] || null;
      snapshot.webusb_button_label = label;
      snapshot.toolpath_calculated = label ? /send file/i.test(label) : false;
      const gate = findOnOffGate(state, outputMod.id);
      if (gate) {
        const cb = gate.params.find(p => p.type === 'checkbox');
        snapshot.output_gate = { on: cb ? cb.value === 'true' : null };
      }
      snapshot.ready_to_send = snapshot.toolpath_calculated
        && snapshot.device_connected
        && (!snapshot.output_gate || snapshot.output_gate.on === true);
    }

    // Size (PNG only — vector formats carry their own)
    if (lastLoadedFile && extname(lastLoadedFile).toLowerCase() === '.png') {
      const info = await browser.getImageInfo();
      if (!info.error) {
        const reader = findReaderModule(state);
        const dpiParam = reader?.params.find(p => p.label.toLowerCase().includes('dpi'));
        const dpi = dpiParam ? parseFloat(dpiParam.value) : null;
        snapshot.size = {
          pixels: `${info.pixelWidth} x ${info.pixelHeight}`,
          ...(dpi ? {
            mm: `${(25.4 * info.pixelWidth / dpi).toFixed(1)} x ${(25.4 * info.pixelHeight / dpi).toFixed(1)}`,
            dpi
          } : {})
        };
      }
    }

    // Next step
    if (!snapshot.file_loaded) snapshot.next_step = 'load_file';
    else if (snapshot.size && !snapshot.size.mm) snapshot.next_step = 'set_physical_size';
    else if (!snapshot.device_connected) snapshot.next_step = 'get_device';
    else if (!snapshot.toolpath_calculated) snapshot.next_step = 'calculate';
    else if (snapshot.output_gate && snapshot.output_gate.on !== true) snapshot.next_step = 'toggle_output_gate';
    else if (snapshot.ready_to_send) snapshot.next_step = 'send_file';
    else snapshot.next_step = 'unknown';

    return composeOk(snapshot);
  }
);

// --- send_job ---

mcpServer.tool('send_job',
  'Composite send-to-machine. Idempotent: callable from any post-setup state — walks back through preconditions (gate, device acquisition, toolpath calculation), auto-fixes what is missing, then sends. Returns {status: "sent", device, send_button_label_after} on confirmed delivery (lower toggle flipped to "waiting for file"), or {status: "failed", at_step, reason, next_action?} when a precondition cannot be auto-recovered. Takes no parameters; toolpath state is read from the lower toggle button.',
  {
    // Accepted but ignored — kept for one version so existing callers passing
    // {skip_calculate: true} don't break. Calculate is now state-driven.
    skip_calculate: z.boolean().optional().describe('DEPRECATED: ignored. send_job now auto-detects toolpath state from the lower toggle.')
  },
  async () => {
    if (!browser.isLaunched()) {
      return composeError('launch_browser', 'Browser not launched.', { next_action: 'launch_browser' });
    }
    if (!loadedProgram) {
      return composeError('load_program', 'No program loaded.', { next_action: 'setup_cut or load_program' });
    }

    let state = await browser.getProgramState();
    const outputMod = findOutputSendModule(state);
    if (!outputMod) {
      return composeError('find_output', 'No WebUSB/WebSerial output module in loaded program — wrong program for this machine?', { program: loadedProgram });
    }

    // The WebUSB output module has three buttons: "Get Device" / "Forget" (top
    // row, static labels) and a lower button that toggles between "send file"
    // (toolpath ready) and "waiting for file" (no toolpath / toolpath consumed).
    // The toggle is the only state-bearing button; identify it by label, never by index.
    const findSendToggle = (mod) => mod?.buttons?.find(b => /send file|waiting for file/i.test(b)) || null;
    const readToggle = async () => findSendToggle(
      (await browser.getProgramState()).find(m => m.id === outputMod.id)
    );

    // Poll toggle until it matches `desired` regex, or timeout. Returns final value.
    // Mods updates the toggle async after a click; fixed sleeps produced false
    // negatives where the data went through but the UI hadn't caught up yet.
    const pollToggleUntil = async (desired, timeoutMs, intervalMs = 250) => {
      const deadline = Date.now() + timeoutMs;
      let last = await readToggle();
      while (!desired.test(last || '') && Date.now() < deadline) {
        await new Promise(r => setTimeout(r, intervalMs));
        last = await readToggle();
      }
      return last;
    };

    // --- Convergent preconditions: walk back, fix what's fixable ---

    // Gate: ensure on (idempotent — skips if already on).
    const gate = findOnOffGate(state, outputMod.id);
    if (gate) {
      const gateCheckbox = gate.params.find(p => p.type === 'checkbox');
      if (gateCheckbox && gateCheckbox.value !== 'true') {
        const r = await browser.setModuleInput(gate.id, gateCheckbox.label, 'true');
        if (r.error) return composeError('toggle_gate', r.error, { gate_id: gate.id });
      }
    }

    // Device: ensure acquired (USBDevice.opened === true), not just granted.
    // Persistent profile keeps grants forever; opened reflects active session.
    let granted = await browser.getGrantedDevices();
    let acquired = granted.find(d => d.opened === true);
    if (!acquired) {
      const r = await browser.clickModuleButton(outputMod.id, 'get device');
      if (r.error) return composeError('click_get_device', r.error);
      // CDP auto-selects from the picker; give it a beat to resolve and mods to open().
      await new Promise(res => setTimeout(res, 1500));
      granted = await browser.getGrantedDevices();
      acquired = granted.find(d => d.opened === true);
    }
    if (!acquired) {
      return composeError('verify_device', 'No device acquired after Get Device click. Check cable + power + that the machine is the granted one.', {
        granted_devices: granted.map(d => d.name),
        next_action: 'verify cable and power, then retry send_job'
      });
    }
    const device = acquired.name;

    // Toolpath: read toggle to decide whether calculate is needed.
    // "send file" → toolpath ready, skip to send. "waiting for file" or missing → calculate.
    let toggle = findSendToggle(state.find(m => m.id === outputMod.id));
    if (!/send file/i.test(toggle || '')) {
      const calcMod = findCalculateModule(state);
      if (!calcMod) {
        return composeError('find_calculator', 'No "calculate" button found in program. Source file may not be loaded.', {
          next_action: 'load_file or setup_cut'
        });
      }
      browser.clearDownloads();
      const r = await browser.clickModuleButton(calcMod.id, 'calculate');
      if (r.error) return composeError('calculate_click', r.error);
      await browser.waitForProcessingSignal({ timeout: 30000 });

      // Poll for toggle to flip to "send file" — gives mods up to 3s beyond
      // the processing signal to update its UI before we read it.
      toggle = await pollToggleUntil(/send file/i, 3000);
      if (!/send file/i.test(toggle || '')) {
        return composeError('verify_send_ready', `Toolpath not ready after calculate. Toggle reads "${toggle || 'none'}" (expected "send file"). Source file may not be loaded, or calculation produced no output.`, {
          send_button_label: toggle,
          next_action: 'check file is loaded and program inputs are valid'
        });
      }
    }

    // Click send.
    const clickResult = await browser.clickModuleButton(outputMod.id, 'send file');
    if (clickResult.error) return composeError('click_send_file', clickResult.error);

    // Verify consumption — poll up to 5s for toggle to flip to "waiting for file".
    // Real WebUSB transfers regularly take longer than 800ms; the previous
    // fixed sleep produced false negatives where the cutter actually started
    // cutting but the toggle hadn't updated when we read it.
    const toggleAfter = await pollToggleUntil(/waiting for file/i, 5000);
    if (!/waiting for file/i.test(toggleAfter || '')) {
      return composeError('verify_consumed', `Send did not consume toolpath within 5s. Toggle reads "${toggleAfter || 'unknown'}" (expected "waiting for file"). Cutter may not have received data.`, {
        device,
        send_button_label_after: toggleAfter,
        next_action: 'check cutter power, USB cable, and that machine is ready to receive'
      });
    }

    return composeOk({
      status: 'sent',
      device,
      send_button_label_after: toggleAfter
    });
  }
);

// --- setup_cut ---

mcpServer.tool('setup_cut',
  'Composite setup: find machine, load program, load file, set physical size, set parameters. Parameter names can be generic (e.g. "speed") — mops maps them to the right module via machine-params.json, falling back to substring search.',
  {
    machine_hint: z.string().describe('Machine name, model, or type (e.g., "Roland GX-24", "vinyl cutter")'),
    file_path: z.string().describe('Absolute path to the file to load (PNG or SVG)'),
    size: z.object({
      width: z.number(),
      height: z.number().optional(),
      unit: z.enum(['mm', 'cm', 'in']).default('mm')
    }).optional().describe('Physical size (PNG only)'),
    parameters: z.record(z.string(), z.union([z.string(), z.number()])).optional().describe('Generic parameter map (e.g., {"speed": 20}). Falls back to label substring search.')
  },
  async ({ machine_hint, file_path, size, parameters }) => {
    if (!browser.isLaunched()) return composeError('precondition', 'Browser not launched');

    try { await stat(file_path); }
    catch {
      const details = await buildMissingFileContext(file_path);
      return composeError('load_file', `File not found: ${file_path}`, { details });
    }

    const ext = extname(file_path).toLowerCase();
    if (ext !== '.png' && ext !== '.svg') {
      return composeError('load_file', `setup_cut supports PNG/SVG. For ${ext}, use load_file manually.`);
    }

    // Step 1: match machine
    const profile = await loadProfile();
    if (!profile.machines?.length) return composeError('find_machine', 'No machines in profile. Use update_profile first.');
    const machine = matchProfileMachine(profile, machine_hint);
    if (!machine) return composeError('find_machine', `No machine matching "${machine_hint}" in profile`);

    // Step 2: resolve program path
    let programPath = machine.program;
    if (!programPath) {
      const programs = await getProgramsManifest();
      const keywords = machine.name.toLowerCase().split(/\s+/).filter(w => w.length > 2);
      const match = programs.find(p => {
        const bag = `${p.path} ${p.name || ''}`.toLowerCase();
        return keywords.some(kw => bag.includes(kw));
      });
      if (!match) return composeError('find_program', `No program found for machine "${machine.name}"`);
      programPath = match.path;
    }

    // Step 3: load program
    try {
      await browser.loadProgram(modsUrl, programPath);
      loadedProgram = programPath;
      lastLoadedFile = null;
    } catch (err) {
      return composeError('load_program', err.message, { program: programPath });
    }

    // Step 4: load file (PNG/SVG via postMessage)
    const loadRes = await browser.postMessageFile(file_path);
    if (loadRes.error) return composeError('load_file', loadRes.error);
    lastLoadedFile = file_path;

    const result = { status: 'ready', machine: machine.name, program: programPath, file: file_path };

    // Step 5: physical size (PNG only)
    let state = await browser.getProgramState();
    if (size && ext === '.png') {
      const sizeRes = await applyPhysicalSize(state, file_path, size);
      if (sizeRes.error) return composeError('set_size', sizeRes.error);
      result.size = sizeRes;
      state = await browser.getProgramState();
    }

    // Step 6: parameters (generic names → machine-specific via map, else substring)
    if (parameters && Object.keys(parameters).length > 0) {
      const paramMap = (await loadMachineParams()).machines?.[machine.name]?.params || {};
      const paramResults = {};
      for (const [key, value] of Object.entries(parameters)) {
        const mapping = paramMap[key.toLowerCase()] || paramMap[key];
        let targetMod = null, paramLabel = null;
        if (mapping) {
          targetMod = state.find(m => m.name.toLowerCase().includes(mapping.module.toLowerCase()));
          paramLabel = mapping.parameter;
        }
        if (!targetMod) {
          // Fallback: any module with a param label containing the key
          targetMod = state.find(m => m.params.some(p => p.label.toLowerCase().includes(key.toLowerCase())));
          paramLabel = key;
        }
        if (!targetMod) { paramResults[key] = { error: `No module with parameter matching "${key}"` }; continue; }
        const r = await browser.setModuleInput(targetMod.id, paramLabel, String(value));
        paramResults[key] = { module: targetMod.name, parameter: paramLabel, ...r };
      }
      result.parameters_set = paramResults;
    }

    const snapshot = await buildProgramSnapshot();
    return composeOk({ ...result, ...snapshot });
  }
);

// --- Startup ---
async function start() {
  try {
    await getModulesManifest();
    console.error(`[mops] Connected to ${modsUrl} (${modulesManifest.length} modules)`);
  } catch (err) {
    console.error(`[mops] Error: Cannot reach mods at ${modsUrl}: ${err.message}`);
    process.exit(1);
  }

  console.error(`[mops] Browser will launch on demand (use launch_browser tool)`);

  const transport = new StdioServerTransport();
  await mcpServer.connect(transport);
  console.error('[mops] MCP server running on stdio');
}

let cleaningUp = false;
async function cleanup() {
  if (cleaningUp) return;
  cleaningUp = true;
  console.error('[mops] Shutting down...');
  try { await browser.close(); } catch (err) { console.error(`[mops] browser.close error: ${err.message}`); }
  process.exit(0);
}

process.on('SIGINT', cleanup);
process.on('SIGTERM', cleanup);
// Claude Desktop often closes stdin without signaling — treat EOF as shutdown
// so Chrome exits cleanly and doesn't leave the profile flagged as crashed.
process.stdin.on('end', cleanup);
process.stdin.on('close', cleanup);

start().catch(err => {
  console.error(`[mops] Fatal error: ${err.message}`);
  process.exit(1);
});
