"use strict";

(() => {
  const MAX_PAGES = 250;
  const PREVIEW_SCALE = 0.32;
  const SLIDE_WIDTH = 13.333333;
  const SLIDE_HEIGHT = 7.5;

  const elements = {
    pptToTextButton: document.querySelector("#ppt-to-text-button"),
    dropZone: document.querySelector("#drop-zone"),
    fileInput: document.querySelector("#file-input"),
    browseButton: document.querySelector("#browse-button"),
    status: document.querySelector("#status-message"),
    filePanel: document.querySelector("#file-panel"),
    fileName: document.querySelector("#file-name"),
    fileMeta: document.querySelector("#file-meta"),
    pageCount: document.querySelector("#page-count"),
    previewGrid: document.querySelector("#preview-grid"),
    removeButton: document.querySelector("#remove-file"),
    convertButton: document.querySelector("#convert-button"),
    progressArea: document.querySelector("#progress-area"),
    progressBar: document.querySelector("#progress-bar"),
    progressLabel: document.querySelector("#progress-label"),
  };

  elements.pptToTextButton.addEventListener("click", () => {
    window.location.href = "ppt-to-txt.html";
  });

  let selectedFile = null;
  let pdfDocument = null;
  let pageImages = [];
  let isLoading = false;
  let isConverting = false;

  if (!window.pdfjsLib || !window.PptxGenJS) {
    showError("The converter libraries could not be loaded. Check your internet connection and refresh the page.");
    elements.dropZone.setAttribute("aria-disabled", "true");
    elements.dropZone.classList.add("disabled");
    return;
  }

  window.pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

  elements.browseButton.addEventListener("click", (event) => {
    event.stopPropagation();
    elements.fileInput.click();
  });
  elements.dropZone.addEventListener("click", (event) => {
    if (event.target !== elements.browseButton && !elements.dropZone.classList.contains("disabled")) {
      elements.fileInput.click();
    }
  });
  elements.dropZone.addEventListener("keydown", (event) => {
    if ((event.key === "Enter" || event.key === " ") && !elements.dropZone.classList.contains("disabled")) {
      event.preventDefault();
      elements.fileInput.click();
    }
  });
  elements.fileInput.addEventListener("change", () => {
    if (elements.fileInput.files.length) loadFile(elements.fileInput.files[0]);
  });
  elements.removeButton.addEventListener("click", clearFile);
  elements.convertButton.addEventListener("click", convertFile);

  ["dragenter", "dragover"].forEach((eventName) => {
    elements.dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      event.stopPropagation();
      if (!elements.dropZone.classList.contains("disabled")) elements.dropZone.classList.add("dragover");
    });
  });
  ["dragleave", "drop"].forEach((eventName) => {
    elements.dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      event.stopPropagation();
      elements.dropZone.classList.remove("dragover");
    });
  });
  elements.dropZone.addEventListener("drop", (event) => {
    if (elements.dropZone.classList.contains("disabled")) return;
    const file = event.dataTransfer.files[0];
    if (file) loadFile(file);
  });

  async function loadFile(file) {
    clearStatus();
    if (isLoading || isConverting) {
      showError("Please wait for the current file to finish processing.");
      return;
    }
    if (!isPdf(file)) {
      showError("That file doesn’t look like a PDF. Please choose a .pdf file.");
      return;
    }
    isLoading = true;
    selectedFile = file;
    pageImages = [];
    elements.fileInput.value = "";
    elements.filePanel.hidden = true;
    elements.previewGrid.replaceChildren();
    elements.dropZone.classList.remove("has-file");
    setProgress(0, "Opening your PDF…");
    elements.progressArea.hidden = false;

    try {
      await discardDocument();
      const data = new Uint8Array(await file.arrayBuffer());
      const loadedDocument = await window.pdfjsLib.getDocument({ data }).promise;
      if (loadedDocument.numPages > MAX_PAGES) {
        const pageCount = loadedDocument.numPages;
        await loadedDocument.destroy();
        selectedFile = null;
        throw new Error(`This PDF has ${pageCount} pages. The maximum supported is ${MAX_PAGES}.`);
      }

      pdfDocument = loadedDocument;
      pageImages = new Array(loadedDocument.numPages);
      elements.fileName.textContent = file.name;
      elements.fileMeta.textContent = `${loadedDocument.numPages} ${loadedDocument.numPages === 1 ? "page" : "pages"} · ${formatBytes(file.size)}`;
      elements.pageCount.textContent = `${loadedDocument.numPages} ${loadedDocument.numPages === 1 ? "slide" : "slides"} will be created`;
      elements.previewGrid.replaceChildren();
      elements.filePanel.hidden = false;
      elements.convertButton.disabled = true;
      elements.progressArea.hidden = false;

      for (let pageNumber = 1; pageNumber <= loadedDocument.numPages; pageNumber += 1) {
        const page = await loadedDocument.getPage(pageNumber);
        const originalViewport = page.getViewport({ scale: 1 });
        const viewport = page.getViewport({ scale: PREVIEW_SCALE });
        const canvas = document.createElement("canvas");
        canvas.width = Math.ceil(viewport.width);
        canvas.height = Math.ceil(viewport.height);
        await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
        const image = canvas.toDataURL("image/jpeg", 0.88);
        pageImages[pageNumber - 1] = {
          pageNumber,
          width: originalViewport.width,
          height: originalViewport.height,
        };
        addPreview(image, pageNumber);
        canvas.width = 0;
        canvas.height = 0;
        page.cleanup();
        setProgress(Math.round((pageNumber / loadedDocument.numPages) * 100), `Preparing page ${pageNumber} of ${loadedDocument.numPages}…`);
      }

      elements.convertButton.disabled = false;
      elements.progressArea.hidden = true;
      elements.dropZone.classList.add("has-file");
    } catch (error) {
      await discardDocument();
      elements.filePanel.hidden = true;
      selectedFile = null;
      showError(formatPdfError(error));
    } finally {
      isLoading = false;
      if (!selectedFile) elements.progressArea.hidden = true;
    }
  }

  function addPreview(imageData, pageNumber) {
    const figure = document.createElement("figure");
    figure.className = "preview-item";
    const imageWrap = document.createElement("div");
    imageWrap.className = "preview-image-wrap";
    const image = document.createElement("img");
    image.src = imageData;
    image.alt = `Preview of page ${pageNumber}`;
    imageWrap.append(image);
    const caption = document.createElement("figcaption");
    caption.innerHTML = `<span>Slide ${pageNumber}</span><span>${pageNumber} / ${pageImages.length}</span>`;
    figure.append(imageWrap, caption);
    elements.previewGrid.append(figure);
  }

  async function convertFile() {
    if (!selectedFile || !pageImages.length || isConverting) return;
    if (typeof window.PptxGenJS !== "function") {
      showError("The PowerPoint library is unavailable. Refresh the page and try again.");
      return;
    }

    isConverting = true;
    elements.convertButton.disabled = true;
    elements.progressArea.hidden = false;
    setProgress(0, "Building your presentation…");
    clearStatus();

    try {
      const presentation = new window.PptxGenJS();
      presentation.layout = "LAYOUT_WIDE";
      presentation.author = "Joemar C. Resmondo";
      presentation.subject = "Converted from PDF";
      presentation.title = selectedFile.name.replace(/\.pdf$/i, "");
      presentation.company = "Free No Registration";

      for (let index = 0; index < pageImages.length; index += 1) {
        const pageImage = pageImages[index];
        const page = await pdfDocument.getPage(pageImage.pageNumber);
        const outputScale = Math.min(2.5, 1920 / pageImage.width);
        const viewport = page.getViewport({ scale: outputScale });
        const canvas = document.createElement("canvas");
        canvas.width = Math.ceil(viewport.width);
        canvas.height = Math.ceil(viewport.height);
        await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
        const imageData = canvas.toDataURL("image/jpeg", 0.92);
        const slide = presentation.addSlide();
        const pageRatio = pageImage.width / pageImage.height;
        const slideRatio = SLIDE_WIDTH / SLIDE_HEIGHT;
        let width;
        let height;
        let x;
        let y;

        if (pageRatio > slideRatio) {
          width = SLIDE_WIDTH;
          height = width / pageRatio;
          x = 0;
          y = (SLIDE_HEIGHT - height) / 2;
        } else {
          height = SLIDE_HEIGHT;
          width = height * pageRatio;
          x = (SLIDE_WIDTH - width) / 2;
          y = 0;
        }

        slide.addImage({ data: imageData, x, y, w: width, h: height });
        canvas.width = 0;
        canvas.height = 0;
        page.cleanup();
        setProgress(Math.round(((index + 1) / pageImages.length) * 80), `Adding slide ${index + 1} of ${pageImages.length}…`);
      }

      setProgress(90, "Saving your PowerPoint…");
      const output = await presentation.write({ outputType: "blob" });
      downloadBlob(output, `${selectedFile.name.replace(/\.pdf$/i, "") || "presentation"}.pptx`);
      setProgress(100, "Your presentation is ready!");
      showSuccess("Your PowerPoint is ready and has been downloaded.");
    } catch (error) {
      console.error("PowerPoint conversion failed:", error);
      showError("We couldn’t create the PowerPoint file. Please try again with a different PDF.");
    } finally {
      isConverting = false;
      elements.convertButton.disabled = false;
      window.setTimeout(() => {
        elements.progressArea.hidden = true;
      }, 1600);
    }
  }

  function isPdf(file) {
    return file.type === "application/pdf" || /\.pdf$/i.test(file.name);
  }

  function formatBytes(bytes) {
    if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  function formatPdfError(error) {
    if (error instanceof Error && error.message.startsWith("This PDF has")) return error.message;
    console.error("PDF loading failed:", error);
    return "We couldn’t open this PDF. It may be damaged, encrypted, or unsupported.";
  }

  function setProgress(percent, label) {
    elements.progressBar.style.width = `${percent}%`;
    elements.progressLabel.textContent = label;
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.append(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  async function discardDocument() {
    if (!pdfDocument) return;
    const oldDocument = pdfDocument;
    pdfDocument = null;
    await oldDocument.destroy();
  }

  function clearFile() {
    if (isLoading || isConverting) return;
    void discardDocument();
    selectedFile = null;
    pageImages = [];
    elements.fileInput.value = "";
    elements.previewGrid.replaceChildren();
    elements.filePanel.hidden = true;
    elements.progressArea.hidden = true;
    elements.dropZone.classList.remove("has-file");
    clearStatus();
  }

  function showError(message) {
    elements.status.textContent = message;
    elements.status.className = "status-message";
    elements.status.hidden = false;
  }

  function showSuccess(message) {
    elements.status.textContent = message;
    elements.status.className = "status-message success";
    elements.status.hidden = false;
  }

  function clearStatus() {
    elements.status.hidden = true;
    elements.status.textContent = "";
    elements.status.className = "status-message";
  }
})();
