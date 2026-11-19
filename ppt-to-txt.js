"use strict";

(() => {
  const elements = {
    dropZone: document.querySelector("#drop-zone"),
    fileInput: document.querySelector("#file-input"),
    browseButton: document.querySelector("#browse-button"),
    status: document.querySelector("#status-message"),
    filePanel: document.querySelector("#file-panel"),
    fileName: document.querySelector("#file-name"),
    fileMeta: document.querySelector("#file-meta"),
    slideCount: document.querySelector("#slide-count"),
    textPreview: document.querySelector("#text-preview"),
    removeButton: document.querySelector("#remove-file"),
    copyButton: document.querySelector("#copy-button"),
    downloadButton: document.querySelector("#convert-button"),
  };

  let selectedFile = null;
  let extractedText = "";
  let isLoading = false;

  if (!window.JSZip || !window.CFB) {
    showError("The PowerPoint reader libraries could not be loaded. Check your internet connection and refresh the page.");
    elements.dropZone.classList.add("disabled");
    elements.dropZone.setAttribute("aria-disabled", "true");
    return;
  }

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
  elements.copyButton.addEventListener("click", copyText);
  elements.downloadButton.addEventListener("click", downloadText);

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
    if (isLoading) {
      showError("Please wait for the current presentation to finish processing.");
      return;
    }
    if (!isPowerPoint(file)) {
      showError("That file doesn’t look like a PowerPoint. Please choose a .ppt or .pptx file.");
      return;
    }
    isLoading = true;
    selectedFile = file;
    extractedText = "";
    elements.fileInput.value = "";
    elements.filePanel.hidden = true;
    elements.textPreview.textContent = "";
    try {
      const buffer = await file.arrayBuffer();
      const result = /\.pptx$/i.test(file.name)
        ? await extractPptxText(buffer)
        : extractLegacyPptText(buffer);
      if (!result.text.trim()) {
        throw new Error("No readable slide text was found. The presentation may contain only images or unsupported embedded content.");
      }

      extractedText = result.text;
      elements.fileName.textContent = file.name;
      elements.fileMeta.textContent = `${result.slideCount} ${result.slideCount === 1 ? "slide" : "slides"} · ${formatBytes(file.size)}`;
      elements.slideCount.textContent = `Text from ${result.slideCount} ${result.slideCount === 1 ? "slide" : "slides"}`;
      elements.textPreview.textContent = result.text;
      elements.filePanel.hidden = false;
      elements.dropZone.classList.add("has-file");
      if (result.legacy) showSuccess("Legacy .ppt text was extracted. Slide boundaries and ordering may vary; check the preview before saving.");
    } catch (error) {
      console.error("PowerPoint text extraction failed:", error);
      selectedFile = null;
      showError(error instanceof Error ? error.message : "We couldn’t read this presentation. Try saving it as a .pptx file and upload it again.");
    } finally {
      isLoading = false;
    }
  }

  async function extractPptxText(buffer) {
    const archive = await window.JSZip.loadAsync(buffer);
    const slideEntries = Object.keys(archive.files)
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
      .sort((left, right) => slideNumber(left) - slideNumber(right));
    if (!slideEntries.length) {
      throw new Error("No slides were found in this .pptx file. It may be damaged or password-protected.");
    }

    const orderedEntries = await getPresentationOrder(archive, slideEntries);
    const sections = [];
    for (let index = 0; index < orderedEntries.length; index += 1) {
      const xmlText = await archive.file(orderedEntries[index]).async("string");
      const xml = new DOMParser().parseFromString(xmlText, "application/xml");
      if (xml.querySelector("parsererror")) throw new Error("This .pptx file contains invalid slide data.");
      const paragraphs = [];
      const shapeTree = xml.getElementsByTagNameNS("*", "spTree")[0];
      if (!shapeTree) throw new Error(`Slide ${index + 1} has no readable shape tree.`);
      Array.from(shapeTree.children).forEach((shape) => {
        Array.from(shape.getElementsByTagNameNS("*", "p")).forEach((paragraph) => {
          const text = Array.from(paragraph.getElementsByTagNameNS("*", "t"))
            .map((node) => node.textContent || "")
            .join("");
          if (text.trim()) paragraphs.push(text);
        });
      });
      if (paragraphs.length) sections.push(`--- Slide ${index + 1} ---\n${paragraphs.join("\n")}`);
      else sections.push(`--- Slide ${index + 1} ---`);
    }
    return { text: sections.join("\n\n").trim(), slideCount: orderedEntries.length, legacy: false };
  }

  async function getPresentationOrder(archive, fallbackEntries) {
    const presentationFile = archive.file("ppt/presentation.xml");
    const relationshipsFile = archive.file("ppt/_rels/presentation.xml.rels");
    if (!presentationFile || !relationshipsFile) return fallbackEntries;

    const [presentationText, relationshipsText] = await Promise.all([
      presentationFile.async("string"),
      relationshipsFile.async("string"),
    ]);
    const presentation = new DOMParser().parseFromString(presentationText, "application/xml");
    const relationships = new DOMParser().parseFromString(relationshipsText, "application/xml");
    if (presentation.querySelector("parsererror") || relationships.querySelector("parsererror")) {
      return fallbackEntries;
    }
    const targets = new Map(
      Array.from(relationships.getElementsByTagNameNS("*", "Relationship")).map((relationship) => [
        relationship.getAttribute("Id"),
        normalizeArchivePath(`ppt/${relationship.getAttribute("Target") || ""}`),
      ]),
    );
    const ordered = Array.from(presentation.getElementsByTagNameNS("*", "sldId"))
      .map((slide) => targets.get(getRelationshipId(slide)))
      .filter((path) => path && archive.file(path));
    return ordered.length === fallbackEntries.length ? ordered : fallbackEntries;
  }

  function extractLegacyPptText(buffer) {
    let compound;
    try {
      compound = window.CFB.read(new Uint8Array(buffer), { type: "array" });
    } catch (error) {
      throw new Error("We couldn’t open this .ppt file. It may be damaged or password-protected.");
    }
    const documentStream = window.CFB.find(compound, "PowerPoint Document");
    if (!documentStream || !documentStream.content) {
      throw new Error("The .ppt file has no readable PowerPoint document stream. Try saving it as a .pptx file.");
    }

    const bytes = Uint8Array.from(documentStream.content);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const decoder = new TextDecoder("windows-1252");
    const textRuns = [];
    let offset = 0;
    while (offset + 8 <= bytes.length) {
      const recordType = view.getUint16(offset + 2, true);
      const length = view.getUint32(offset + 4, true);
      const contentStart = offset + 8;
      const contentEnd = contentStart + length;
      if (contentEnd <= bytes.length && (recordType === 4000 || recordType === 4008) && length > 0) {
        let text;
        if (recordType === 4000 && length % 2 === 0) {
          text = new TextDecoder("utf-16le").decode(bytes.subarray(contentStart, contentEnd));
        } else if (recordType === 4008) {
          text = decoder.decode(bytes.subarray(contentStart, contentEnd));
        }
        if (text) {
          text = cleanLegacyText(text);
          if (text.length && hasReadableCharacters(text)) textRuns.push(text);
          offset = contentEnd;
          continue;
        }
      }
      offset += 2;
    }

    const uniqueRuns = textRuns.filter((text, index) => textRuns.indexOf(text) === index);
    if (!uniqueRuns.length) {
      throw new Error("No readable text was found in this legacy .ppt file. Try saving it as a .pptx file.");
    }
    return {
      text: uniqueRuns.join("\n\n"),
      slideCount: Math.max(1, countLegacySlides(bytes, view)),
      legacy: true,
    };
  }

  function countLegacySlides(bytes, view) {
    let count = 0;
    for (let offset = 0; offset + 8 <= bytes.length; offset += 2) {
      if (view.getUint16(offset + 2, true) !== 1011) continue;
      const length = view.getUint32(offset + 4, true);
      if (length >= 16 && offset + 8 + length <= bytes.length) count += 1;
    }
    return count;
  }

  function cleanLegacyText(text) {
    return text
      .replace(/\u0000/g, "")
      .replace(/[\u0001-\u0008\u000B\u000C\u000E-\u001F]/g, "")
      .replace(/\r\n?/g, "\n")
      .trim();
  }

  function hasReadableCharacters(text) {
    const readable = text.match(/[\p{L}\p{N}\p{P}\p{Zs}]/gu);
    return readable !== null && readable.length / text.length >= 0.7;
  }

  function slideNumber(path) {
    return Number(path.match(/slide(\d+)\.xml$/i)?.[1] || 0);
  }

  function getRelationshipId(element) {
    const relationshipId = element.getAttributeNS(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "id",
    );
    if (relationshipId) return relationshipId;
    const attribute = Array.from(element.attributes).find((item) => item.localName === "id" && item.name !== "id");
    return attribute ? attribute.value : null;
  }

  function normalizeArchivePath(path) {
    const parts = [];
    path.split("/").forEach((part) => {
      if (part === "..") parts.pop();
      else if (part && part !== ".") parts.push(part);
    });
    return parts.join("/");
  }

  function isPowerPoint(file) {
    return /\.pptx?$/i.test(file.name) ||
      file.type === "application/vnd.ms-powerpoint" ||
      file.type === "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  }

  function formatBytes(bytes) {
    if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  async function copyText() {
    if (!extractedText) return;
    try {
      await navigator.clipboard.writeText(extractedText);
      showSuccess("Slide text copied to your clipboard.");
    } catch (error) {
      console.error("Clipboard copy failed:", error);
      showError("We couldn’t access the clipboard. Select and copy the text from the preview instead.");
    }
  }

  function downloadText() {
    if (!selectedFile || !extractedText) return;
    const filename = selectedFile.name.replace(/\.pptx?$/i, "") || "presentation";
    downloadBlob(new Blob([extractedText], { type: "text/plain;charset=utf-8" }), `${filename}.txt`);
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

  function clearFile() {
    if (isLoading) return;
    selectedFile = null;
    extractedText = "";
    elements.fileInput.value = "";
    elements.textPreview.textContent = "";
    elements.filePanel.hidden = true;
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
