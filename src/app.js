/* global Office, PowerPoint, console */

// Last copied shape position, kept in memory for the life of the shared runtime.
let copiedPosition = null;

Office.onReady(() => {
  // No-op: Office.actions.associate calls below register the shortcut handlers.
});

function showStatus(message) {
  const el = document.getElementById("status");
  if (el) {
    el.textContent = message;
  }
  console.log(message);
}

async function withSelectedShapes(minCount, callback) {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();

      shapes.items.forEach((shape) => shape.load("left,top,width,height,id"));
      await context.sync();

      if (shapes.items.length < minCount) {
        showStatus(`Select at least ${minCount} shape(s) first.`);
        return;
      }

      await callback(context, shapes.items);
      await context.sync();
    });
  } catch (err) {
    showStatus(`Error: ${err.message}`);
    console.error(err);
  }
}

function selectionBounds(items) {
  const left = Math.min(...items.map((s) => s.left));
  const right = Math.max(...items.map((s) => s.left + s.width));
  const top = Math.min(...items.map((s) => s.top));
  const bottom = Math.max(...items.map((s) => s.top + s.height));
  return { left, right, top, bottom };
}

async function readClipboardText() {
  if (navigator.clipboard && navigator.clipboard.readText) {
    try {
      return await navigator.clipboard.readText();
    } catch (err) {
      console.warn("navigator.clipboard.readText failed:", err);
    }
  }

  return new Promise((resolve, reject) => {
    try {
      const textArea = document.createElement("textarea");
      textArea.style.position = "fixed";
      textArea.style.left = "-9999px";
      document.body.appendChild(textArea);
      textArea.focus();
      const pasted = document.execCommand("paste");
      const text = textArea.value;
      textArea.remove();
      if (!pasted && !text) {
        reject(new Error("Could not read clipboard."));
        return;
      }
      resolve(text);
    } catch (err) {
      reject(err);
    }
  });
}

function trimTrailingNewlines(text) {
  return text.replace(/[\r\n]+$/, "");
}

function insertPlainTextAtSelection(text) {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      text,
      { coercionType: Office.CoercionType.Text },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error?.message || "Could not insert text."));
        }
      }
    );
  });
}

// ---- Align ----

Office.actions.associate("AlignLeft", () =>
  withSelectedShapes(2, async (context, items) => {
    const left = Math.min(...items.map((s) => s.left));
    items.forEach((s) => (s.left = left));
  })
);

Office.actions.associate("AlignRight", () =>
  withSelectedShapes(2, async (context, items) => {
    const right = Math.max(...items.map((s) => s.left + s.width));
    items.forEach((s) => (s.left = right - s.width));
  })
);

Office.actions.associate("AlignCenter", () =>
  withSelectedShapes(2, async (context, items) => {
    const { left, right } = selectionBounds(items);
    const center = (left + right) / 2;
    items.forEach((s) => (s.left = center - s.width / 2));
  })
);

Office.actions.associate("AlignTop", () =>
  withSelectedShapes(2, async (context, items) => {
    const top = Math.min(...items.map((s) => s.top));
    items.forEach((s) => (s.top = top));
  })
);

Office.actions.associate("AlignBottom", () =>
  withSelectedShapes(2, async (context, items) => {
    const bottom = Math.max(...items.map((s) => s.top + s.height));
    items.forEach((s) => (s.top = bottom - s.height));
  })
);

Office.actions.associate("AlignMiddle", () =>
  withSelectedShapes(2, async (context, items) => {
    const { top, bottom } = selectionBounds(items);
    const middle = (top + bottom) / 2;
    items.forEach((s) => (s.top = middle - s.height / 2));
  })
);

// ---- Distribute ----
// Equal-gap distribution: the outermost two shapes stay put; shapes between
// them are spaced so the gaps between edges are equal.

Office.actions.associate("DistributeHorizontal", () =>
  withSelectedShapes(3, async (context, items) => {
    const sorted = [...items].sort((a, b) => a.left - b.left);
    const first = sorted[0];
    const last = sorted[sorted.length - 1];
    const middle = sorted.slice(1, -1);
    const span = last.left - (first.left + first.width);
    const middleWidths = middle.reduce((sum, s) => sum + s.width, 0);
    const gap = (span - middleWidths) / (middle.length + 1);

    let cursor = first.left + first.width;
    middle.forEach((s) => {
      cursor += gap;
      s.left = cursor;
      cursor += s.width;
    });
  })
);

Office.actions.associate("DistributeVertical", () =>
  withSelectedShapes(3, async (context, items) => {
    const sorted = [...items].sort((a, b) => a.top - b.top);
    const first = sorted[0];
    const last = sorted[sorted.length - 1];
    const middle = sorted.slice(1, -1);
    const span = last.top - (first.top + first.height);
    const middleHeights = middle.reduce((sum, s) => sum + s.height, 0);
    const gap = (span - middleHeights) / (middle.length + 1);

    let cursor = first.top + first.height;
    middle.forEach((s) => {
      cursor += gap;
      s.top = cursor;
      cursor += s.height;
    });
  })
);

// ---- Resize ----
// Reference shape is the first item PowerPoint returns for the selection
// (the JS API does not expose "last selected" the way the desktop UI does).

Office.actions.associate("SameSize", () =>
  withSelectedShapes(2, async (context, items) => {
    const ref = items[0];
    items.slice(1).forEach((s) => {
      s.width = ref.width;
      s.height = ref.height;
    });
  })
);

Office.actions.associate("SameWidth", () =>
  withSelectedShapes(2, async (context, items) => {
    const ref = items[0];
    items.slice(1).forEach((s) => (s.width = ref.width));
  })
);

Office.actions.associate("SameHeight", () =>
  withSelectedShapes(2, async (context, items) => {
    const ref = items[0];
    items.slice(1).forEach((s) => (s.height = ref.height));
  })
);

// ---- Arrange (z-order) ----

function changeZOrder(operation) {
  return withSelectedShapes(1, async (context, items) => {
    items.forEach((s) => s.setZOrder(operation));
  });
}

Office.actions.associate("BringForward", () => changeZOrder("BringForward"));
Office.actions.associate("BringToFront", () => changeZOrder("BringToFront"));
Office.actions.associate("SendBackward", () => changeZOrder("SendBackward"));
Office.actions.associate("SendToBack", () => changeZOrder("SendToBack"));

// ---- Copy / Paste position ----

Office.actions.associate("CopyPosition", () =>
  withSelectedShapes(1, async (context, items) => {
    const s = items[0];
    copiedPosition = { left: s.left, top: s.top, width: s.width, height: s.height };
    showStatus("Position copied.");
  })
);

Office.actions.associate("PastePosition", () =>
  withSelectedShapes(1, async (context, items) => {
    if (!copiedPosition) {
      showStatus("Nothing copied yet. Use Copy Position first.");
      return;
    }
    items.forEach((s) => {
      s.left = copiedPosition.left;
      s.top = copiedPosition.top;
      s.width = copiedPosition.width;
      s.height = copiedPosition.height;
    });
  })
);

Office.actions.associate("SwapPositions", () =>
  withSelectedShapes(2, async (context, items) => {
    if (items.length !== 2) {
      showStatus("Select exactly 2 shapes to swap positions.");
      return;
    }
    const [first, second] = items;
    const firstLeft = first.left;
    const firstTop = first.top;
    first.left = second.left;
    first.top = second.top;
    second.left = firstLeft;
    second.top = firstTop;
    showStatus("Positions swapped.");
  })
);

// ---- Insert: Sticky Note ----

Office.actions.associate("InsertStickyNote", () => {
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const selectedShapes = context.presentation.getSelectedShapes();
    selectedShapes.load("items");
    await context.sync();
    selectedShapes.items.forEach((s) => s.load("left,top,width"));
    await context.sync();

    let left = 80;
    let top = 80;
    if (selectedShapes.items.length > 0) {
      const ref = selectedShapes.items[0];
      left = ref.left + ref.width + 16;
      top = ref.top;
    }

    const note = slide.shapes.addTextBox("Note", {
      left,
      top,
      width: 160,
      height: 90,
    });
    note.name = "Sticky Note";
    note.fill.setSolidColor("#FFF2A8");
    note.lineFormat.color = "#E0C93A";
    note.lineFormat.weight = 1;
    note.textFrame.textRange.font.size = 11;
    note.textFrame.textRange.font.color = "#5C4B00";
    await context.sync();

    note.select();
    await context.sync();
  }).catch((err) => {
    showStatus(`Error: ${err.message}`);
    console.error(err);
  });
});

// ---- Insert: Text Box ----

Office.actions.associate("InsertTextBox", () => {
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const selectedShapes = context.presentation.getSelectedShapes();
    selectedShapes.load("items");
    await context.sync();
    selectedShapes.items.forEach((s) => s.load("left,top,width"));
    await context.sync();

    let left = 80;
    let top = 80;
    if (selectedShapes.items.length > 0) {
      const ref = selectedShapes.items[0];
      left = ref.left + ref.width + 16;
      top = ref.top;
    }

    const textBox = slide.shapes.addTextBox("Text", { left, top, width: 200, height: 40 });
    textBox.name = "Text Box";

    const textFrame = textBox.textFrame;
    textFrame.leftMargin = 0;
    textFrame.rightMargin = 0;
    textFrame.topMargin = 0;
    textFrame.bottomMargin = 0;
    textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeShapeToFitText;
    textFrame.wordWrap = true;

    textBox.lineFormat.visible = false;
    textBox.lineFormat.transparency = 1;
    textBox.lineFormat.weight = 0;
    textBox.fill.clear();

    await context.sync();
    textBox.select();
    await context.sync();
  }).catch((err) => {
    showStatus(`Error: ${err.message}`);
    console.error(err);
  });
});

// ---- Text: Wrap Text in Shape ----

Office.actions.associate("ToggleWrapText", () =>
  withSelectedShapes(1, async (context, items) => {
    const shape = items[0];
    shape.textFrame.load("wordWrap");
    await context.sync();

    shape.textFrame.wordWrap = !shape.textFrame.wordWrap;
    showStatus(shape.textFrame.wordWrap ? "Wrap text on." : "Wrap text off.");
  })
);

// ---- Text: Toggle Bullets ----

Office.actions.associate("ToggleBullets", () =>
  withSelectedShapes(1, async (context, items) => {
    const shape = items[0];
    shape.load("textFrame");
    await context.sync();

    const textRange = shape.textFrame.textRange;
    textRange.load("paragraphFormat/bulletFormat/visible");
    await context.sync();

    const bulletFormat = textRange.paragraphFormat.bulletFormat;
    const currentlyVisible = bulletFormat.visible === true;
    bulletFormat.visible = !currentlyVisible;
    if (!currentlyVisible) {
      bulletFormat.type = PowerPoint.BulletType.unnumbered;
    }
    showStatus(!currentlyVisible ? "Bullets on." : "Bullets off.");
  })
);

// ---- Text: Paste Unformatted ----

Office.actions.associate("PasteUnformatted", async () => {
  try {
    const rawText = await readClipboardText();
    if (!rawText) {
      showStatus("Clipboard is empty.");
      return;
    }

    const text = trimTrailingNewlines(rawText);
    await insertPlainTextAtSelection(text);
    showStatus("Pasted unformatted text.");
  } catch (err) {
    showStatus("Click inside a text box first, then try again.");
    console.error(err);
  }
});
