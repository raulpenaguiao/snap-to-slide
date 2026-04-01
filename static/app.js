// ── DOM refs ──────────────────────────────────────────────────────────────────

const dropZone       = document.getElementById('drop-zone');
const fileInput      = document.getElementById('file-input');

const previewSection = document.getElementById('preview-section');
const previewImg     = document.getElementById('preview-img');
const previewName    = document.getElementById('preview-name');
const rotateLeftBtn  = document.getElementById('rotate-left-btn');
const rotateRightBtn = document.getElementById('rotate-right-btn');
const changeBtn      = document.getElementById('change-btn');

const videoPreviewSection = document.getElementById('video-preview-section');
const previewVideo   = document.getElementById('preview-video');
const frameSlider    = document.getElementById('frame-slider');
const frameTime      = document.getElementById('frame-time');
const useFrameBtn    = document.getElementById('use-frame-btn');
const videoFilename  = document.getElementById('video-filename');

const hintInput      = document.getElementById('hint-input');

const generateBtn    = document.getElementById('generate-btn');
const progressPanel  = document.getElementById('progress-panel');

const contentPreview = document.getElementById('content-preview');
const cpHeader       = document.getElementById('cp-header');
const cpBody         = document.getElementById('cp-body');
const cpSlideTitle   = document.getElementById('cp-slide-title');
const cpItems        = document.getElementById('cp-items');

const downloadBtn    = document.getElementById('download-btn');
const errorBox       = document.getElementById('error-box');
const resetLink      = document.getElementById('reset-link');
const stepItems      = document.querySelectorAll('.step-item');
const apiKeyInput    = document.getElementById('api-key-input');
const toggleKeyBtn   = document.getElementById('toggle-key-btn');

const themeSelect    = document.getElementById('theme-select');
const themeSwatch    = document.getElementById('theme-swatch');
const themeDesc      = document.getElementById('theme-desc');

// ── State ─────────────────────────────────────────────────────────────────────

let selectedFile      = null;   // raw File object (image or video)
let capturedFrameBlob = null;   // JPEG blob captured from video canvas
let isVideoMode       = false;
let currentES         = null;   // active EventSource
let rotationDeg       = 0;      // current preview rotation: 0 | 90 | 180 | 270

// ── Theme selector ────────────────────────────────────────────────────────────

let availableThemes = [];

function updateThemeSwatch() {
  const selected = availableThemes.find(t => t.name === themeSelect.value);
  if (!selected) return;
  themeSwatch.style.background = `#${selected.bg}`;
  themeSwatch.style.borderColor = `#${selected.accent}`;
  themeDesc.textContent = selected.desc + (selected.has_template ? ' ✦' : '');
}

fetch('./themes')
  .then(r => r.json())
  .then(data => {
    availableThemes = data.themes || [];
    themeSelect.innerHTML = '';
    availableThemes.forEach(t => {
      const opt = document.createElement('option');
      opt.value = t.name;
      opt.textContent = t.name;
      themeSelect.appendChild(opt);
    });
    updateThemeSwatch();
  })
  .catch(() => {
    // Server not yet updated — add a single default option so the form still works
    const opt = document.createElement('option');
    opt.value = 'Default';
    opt.textContent = 'Default';
    themeSelect.appendChild(opt);
  });

themeSelect.addEventListener('change', updateThemeSwatch);

// ── API key show/hide ─────────────────────────────────────────────────────────

toggleKeyBtn.addEventListener('click', () => {
  const isPassword = apiKeyInput.type === 'password';
  apiKeyInput.type = isPassword ? 'text' : 'password';
  toggleKeyBtn.textContent = isPassword ? 'Hide' : 'Show';
});

const savedKey = sessionStorage.getItem('anthropic_api_key');
if (savedKey) apiKeyInput.value = savedKey;

apiKeyInput.addEventListener('input', () => {
  sessionStorage.setItem('anthropic_api_key', apiKeyInput.value);
});

// ── Drop zone ─────────────────────────────────────────────────────────────────

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('keydown', e => {
  if (e.key === 'Enter' || e.key === ' ') fileInput.click();
});

dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file && (file.type.startsWith('image/') || file.type.startsWith('video/'))) {
    setFile(file);
  }
});

fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) setFile(fileInput.files[0]);
});

changeBtn.addEventListener('click', () => fileInput.click());

// ── File selection ────────────────────────────────────────────────────────────

function applyRotation() {
  previewImg.style.transform = rotationDeg ? `rotate(${rotationDeg}deg)` : '';
}

rotateLeftBtn.addEventListener('click', () => {
  rotationDeg = (rotationDeg - 90 + 360) % 360;
  applyRotation();
});

rotateRightBtn.addEventListener('click', () => {
  rotationDeg = (rotationDeg + 90) % 360;
  applyRotation();
});

// ── Rotate a blob by degrees using canvas (baked into upload) ─────────────────

function getRotatedBlob(source, degrees) {
  return new Promise(resolve => {
    if (degrees === 0) { resolve(source); return; }
    const img = new Image();
    const url = URL.createObjectURL(source);
    img.onload = () => {
      URL.revokeObjectURL(url);
      const swap = degrees === 90 || degrees === 270;
      const canvas = document.createElement('canvas');
      canvas.width  = swap ? img.naturalHeight : img.naturalWidth;
      canvas.height = swap ? img.naturalWidth  : img.naturalHeight;
      const ctx = canvas.getContext('2d');
      ctx.translate(canvas.width / 2, canvas.height / 2);
      ctx.rotate(degrees * Math.PI / 180);
      ctx.drawImage(img, -img.naturalWidth / 2, -img.naturalHeight / 2);
      canvas.toBlob(blob => resolve(blob || source), 'image/jpeg', 0.92);
    };
    img.onerror = () => { URL.revokeObjectURL(url); resolve(source); };
    img.src = url;
  });
}

function setFile(file) {
  selectedFile = file;
  capturedFrameBlob = null;
  isVideoMode = file.type.startsWith('video/');
  rotationDeg = 0;

  hideError();
  hideContentPreview();
  downloadBtn.classList.remove('visible');
  resetLink.classList.remove('visible');
  progressPanel.classList.remove('visible');
  progressPanel.querySelector('h3').textContent = 'Processing…';
  resetSteps();

  if (isVideoMode) {
    previewSection.classList.remove('visible');
    videoPreviewSection.classList.add('visible');
    videoFilename.textContent = file.name;
    useFrameBtn.disabled = true;
    useFrameBtn.textContent = 'Capture this frame';
    frameTime.textContent = '0.0s';
    previewVideo.src = URL.createObjectURL(file);
    previewVideo.load();
    generateBtn.disabled = true;  // must capture a frame first
  } else {
    videoPreviewSection.classList.remove('visible');
    previewImg.src = URL.createObjectURL(file);
    previewImg.style.transform = '';
    previewName.textContent = file.name;
    previewSection.classList.add('visible');
    generateBtn.disabled = false;
  }
}

// ── Video: metadata loaded → enable slider ────────────────────────────────────

previewVideo.addEventListener('loadedmetadata', () => {
  frameSlider.max = previewVideo.duration;
  frameSlider.value = 0;
  frameTime.textContent = '0.0s';
  useFrameBtn.disabled = false;
});

// ── Video: slider seek ────────────────────────────────────────────────────────

frameSlider.addEventListener('input', () => {
  const t = parseFloat(frameSlider.value);
  previewVideo.currentTime = t;
  frameTime.textContent = t.toFixed(1) + 's';
});

// ── Video: capture frame ──────────────────────────────────────────────────────

useFrameBtn.addEventListener('click', () => {
  const vw = previewVideo.videoWidth;
  const vh = previewVideo.videoHeight;
  if (!vw || !vh) return;

  const canvas = document.createElement('canvas');
  canvas.width = vw;
  canvas.height = vh;
  canvas.getContext('2d').drawImage(previewVideo, 0, 0, vw, vh);

  canvas.toBlob(blob => {
    if (!blob) return;
    capturedFrameBlob = blob;
    rotationDeg = 0;

    previewImg.src = URL.createObjectURL(blob);
    previewImg.style.transform = '';
    const t = parseFloat(frameSlider.value).toFixed(1);
    previewName.textContent = `Frame at ${t}s — ${selectedFile.name}`;
    previewSection.classList.add('visible');
    videoPreviewSection.classList.remove('visible');

    useFrameBtn.textContent = `Frame at ${t}s captured ✓`;
    generateBtn.disabled = false;
  }, 'image/jpeg', 0.92);
});

// ── Progress step helpers ─────────────────────────────────────────────────────

function resetSteps() {
  stepItems.forEach(item => item.classList.remove('active', 'done'));
}

function setStep(stepNum, status) {
  stepItems.forEach(item => {
    if (parseInt(item.dataset.step) === stepNum) {
      item.classList.remove('active', 'done');
      if (status === 'active') item.classList.add('active');
      if (status === 'done')   item.classList.add('done');
    }
  });
}

// ── Error helpers ─────────────────────────────────────────────────────────────

function showError(msg) {
  errorBox.textContent = msg;
  errorBox.classList.add('visible');
  resetLink.classList.add('visible');
}

function hideError() {
  errorBox.classList.remove('visible');
  resetLink.classList.remove('visible');
}

// ── Content preview helpers ───────────────────────────────────────────────────

cpHeader.addEventListener('click', () => {
  const collapsed = cpHeader.classList.toggle('collapsed');
  cpBody.style.display = collapsed ? 'none' : '';
});

function showContentPreview(preview) {
  cpSlideTitle.textContent = preview.title || '';
  cpItems.innerHTML = '';

  for (const item of (preview.items || [])) {
    const div = document.createElement('div');
    div.className = `cp-item kind-${item.kind}`;
    div.textContent = item.text;
    cpItems.appendChild(div);
  }

  cpHeader.classList.remove('collapsed');
  cpBody.style.display = '';
  contentPreview.classList.add('visible');
}

function hideContentPreview() {
  contentPreview.classList.remove('visible');
  cpSlideTitle.textContent = '';
  cpItems.innerHTML = '';
}

// ── Generate ──────────────────────────────────────────────────────────────────

generateBtn.addEventListener('click', async () => {
  const fileToUpload = isVideoMode ? capturedFrameBlob : selectedFile;
  if (!fileToUpload) return;

  const apiKey = apiKeyInput.value.trim();
  if (!apiKey) {
    showError('Please enter your Anthropic API key.');
    apiKeyInput.focus();
    return;
  }

  generateBtn.disabled = true;
  hideError();
  hideContentPreview();
  downloadBtn.classList.remove('visible');
  progressPanel.classList.add('visible');
  resetSteps();

  if (currentES) { currentES.close(); currentES = null; }

  const rotatedBlob = await getRotatedBlob(fileToUpload, rotationDeg);
  const uploadName  = isVideoMode ? 'frame.jpg' : selectedFile.name;
  const formData = new FormData();
  formData.append('file', new File([rotatedBlob], uploadName, { type: 'image/jpeg' }));
  formData.append('api_key', apiKey);

  const hint = hintInput.value.trim();
  if (hint) formData.append('hint_text', hint);
  formData.append('theme', themeSelect.value || 'Default');

  let jobId;
  try {
    const res = await fetch('./generate', { method: 'POST', body: formData });
    if (!res.ok) {
      const err = await res.json().catch(() => ({ detail: 'Upload failed' }));
      throw new Error(err.detail || 'Upload failed');
    }
    jobId = (await res.json()).job_id;
  } catch (err) {
    showError('Upload failed: ' + err.message);
    generateBtn.disabled = false;
    return;
  }

  const es = new EventSource(`./stream/${jobId}`);
  currentES = es;

  es.onmessage = (event) => {
    let payload;
    try { payload = JSON.parse(event.data); } catch { return; }

    if (payload.error) {
      es.close();
      currentES = null;
      showError('Error: ' + payload.error);
      generateBtn.disabled = false;
      return;
    }

    if (payload.step)    setStep(payload.step, payload.status);
    if (payload.preview) showContentPreview(payload.preview);

    if (payload.status === 'done' && payload.download_token) {
      es.close();
      currentES = null;
      const token = payload.download_token;
      downloadBtn.onclick = () => { window.location.href = `./download/${token}`; };
      downloadBtn.classList.add('visible');
      resetLink.classList.add('visible');
      progressPanel.querySelector('h3').textContent = 'Done!';
    }
  };

  es.onerror = () => {
    if (es.readyState === EventSource.CLOSED) return;
    es.close();
    currentES = null;
    if (!downloadBtn.classList.contains('visible')) {
      showError('Connection lost. Please try again.');
      generateBtn.disabled = false;
    }
  };
});

// ── Start over ────────────────────────────────────────────────────────────────

document.getElementById('start-over').addEventListener('click', () => {
  if (currentES) { currentES.close(); currentES = null; }

  selectedFile = null;
  capturedFrameBlob = null;
  isVideoMode = false;
  rotationDeg = 0;

  fileInput.value = '';
  previewImg.src = '';
  previewImg.style.transform = '';
  previewVideo.src = '';
  hintInput.value = '';

  previewSection.classList.remove('visible');
  videoPreviewSection.classList.remove('visible');

  useFrameBtn.textContent = 'Capture this frame';
  useFrameBtn.disabled = true;
  frameSlider.value = 0;
  frameTime.textContent = '0.0s';

  generateBtn.disabled = true;
  progressPanel.classList.remove('visible');
  progressPanel.querySelector('h3').textContent = 'Processing…';
  downloadBtn.classList.remove('visible');

  hideError();
  hideContentPreview();
  resetSteps();
});
