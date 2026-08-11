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

function shapeCenterY(shape) {
  return shape.top + shape.height / 2;
}

function shapeCenterX(shape) {
  return shape.left + shape.width / 2;
}

function averageSize(shapes, dimension) {
  if (shapes.length === 0) {
    return 20;
  }
  const total = shapes.reduce(
    (sum, shape) => sum + (dimension === "height" ? shape.height : shape.width),
    0
  );
  return total / shapes.length;
}

// Cluster shapes that share roughly the same row or column.
function clusterShapes(shapes, getCenter, tolerance) {
  if (shapes.length === 0) {
    return [];
  }

  const sorted = [...shapes].sort((a, b) => getCenter(a) - getCenter(b));
  const clusters = [[sorted[0]]];

  for (let i = 1; i < sorted.length; i++) {
    const cluster = clusters[clusters.length - 1];
    const clusterAvg =
      cluster.reduce((sum, shape) => sum + getCenter(shape), 0) / cluster.length;

    if (Math.abs(getCenter(sorted[i]) - clusterAvg) <= tolerance) {
      cluster.push(sorted[i]);
    } else {
      clusters.push([sorted[i]]);
    }
  }

  return clusters;
}

async function groupSelectedByAxis(axis) {
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      if (selectedShapes.items.length < 2) {
        showStatus("Select at least 2 shapes to group.");
        return;
      }

      selectedShapes.items.forEach((shape) => shape.load("left,top,width,height,id"));
      await context.sync();

      const items = selectedShapes.items;
      const isRow = axis === "row";
      const getCenter = isRow ? shapeCenterY : shapeCenterX;
      const tolerance = averageSize(items, isRow ? "height" : "width") * 0.35;
      const clusters = clusterShapes(items, getCenter, tolerance);
      const groupsToCreate = clusters.filter((cluster) => cluster.length >= 2);

      if (groupsToCreate.length === 0) {
        showStatus(`No ${axis}s found with 2+ shapes.`);
        return;
      }

      groupsToCreate.forEach((cluster) => slide.shapes.addGroup(cluster));
      await context.sync();

      const label = isRow ? "row" : "column";
      showStatus(`Created ${groupsToCreate.length} ${label} group(s).`);
    });
  } catch (err) {
    showStatus(`Error: ${err.message}`);
    console.error(err);
  }
}

async function withAlignShapes(callback) {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();

      if (shapes.items.length === 0) {
        showStatus("Select at least one shape first.");
        return;
      }

      shapes.items.forEach((shape) => shape.load("left,top,width,height,id"));
      await context.sync();

      let slideBounds = null;
      if (shapes.items.length === 1) {
        const pageSetup = context.presentation.pageSetup;
        pageSetup.load("slideWidth,slideHeight");
        await context.sync();
        slideBounds = {
          width: pageSetup.slideWidth,
          height: pageSetup.slideHeight,
        };
      }

      await callback(shapes.items, slideBounds);
      await context.sync();
    });
  } catch (err) {
    showStatus(`Error: ${err.message}`);
    console.error(err);
  }
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
// 1 shape: align to the slide edges/center. 2+ shapes: align to each other.

Office.actions.associate("AlignLeft", () =>
  withAlignShapes(async (items, slideBounds) => {
    const targetLeft = slideBounds ? 0 : Math.min(...items.map((s) => s.left));
    items.forEach((s) => (s.left = targetLeft));
  })
);

Office.actions.associate("AlignRight", () =>
  withAlignShapes(async (items, slideBounds) => {
    if (slideBounds) {
      items.forEach((s) => (s.left = slideBounds.width - s.width));
      return;
    }
    const right = Math.max(...items.map((s) => s.left + s.width));
    items.forEach((s) => (s.left = right - s.width));
  })
);

Office.actions.associate("AlignCenter", () =>
  withAlignShapes(async (items, slideBounds) => {
    if (slideBounds) {
      items.forEach((s) => (s.left = (slideBounds.width - s.width) / 2));
      return;
    }
    const { left, right } = selectionBounds(items);
    const center = (left + right) / 2;
    items.forEach((s) => (s.left = center - s.width / 2));
  })
);

Office.actions.associate("AlignTop", () =>
  withAlignShapes(async (items, slideBounds) => {
    const targetTop = slideBounds ? 0 : Math.min(...items.map((s) => s.top));
    items.forEach((s) => (s.top = targetTop));
  })
);

Office.actions.associate("AlignBottom", () =>
  withAlignShapes(async (items, slideBounds) => {
    if (slideBounds) {
      items.forEach((s) => (s.top = slideBounds.height - s.height));
      return;
    }
    const bottom = Math.max(...items.map((s) => s.top + s.height));
    items.forEach((s) => (s.top = bottom - s.height));
  })
);

Office.actions.associate("AlignMiddle", () =>
  withAlignShapes(async (items, slideBounds) => {
    if (slideBounds) {
      items.forEach((s) => (s.top = (slideBounds.height - s.height) / 2));
      return;
    }
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

// ---- Group by row / column ----

Office.actions.associate("GroupByRow", () => groupSelectedByAxis("row"));
Office.actions.associate("GroupByColumn", () => groupSelectedByAxis("column"));

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
    note.fill.setSolidColor("#FFFF00");
    note.lineFormat.color = "#E0E000";
    note.lineFormat.weight = 1;
    note.textFrame.textRange.font.size = 14;
    note.textFrame.textRange.font.color = "#000000";
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

// ---- Text: Margins ----
// Adjusts all four text frame margins uniformly (in points).
// Uses per-shape session tracking so we don't rely on reading margins back from the API.

const MARGIN_STEP_PT = 6;
const shapeMarginLevels = new Map();

function setUniformMargins(shape, marginPt) {
  const tf = shape.textFrame;
  tf.autoSizeSetting = PowerPoint.ShapeAutoSize.none;
  tf.leftMargin = marginPt;
  tf.rightMargin = marginPt;
  tf.topMargin = marginPt;
  tf.bottomMargin = marginPt;
}

function adjustTextMargins(delta) {
  return PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();

    if (shapes.items.length === 0) {
      showStatus("Select a shape with text first.");
      return;
    }

    shapes.items.forEach((shape) => shape.load("id"));
    await context.sync();

    const needsSeed = shapes.items.some((shape) => !shapeMarginLevels.has(shape.id));
    if (needsSeed) {
      shapes.items.forEach((shape) => {
        if (!shapeMarginLevels.has(shape.id)) {
          shape.textFrame.load("leftMargin");
        }
      });
      await context.sync();
      shapes.items.forEach((shape) => {
        if (!shapeMarginLevels.has(shape.id)) {
          shapeMarginLevels.set(shape.id, shape.textFrame.leftMargin ?? 0);
        }
      });
    }

    shapes.items.forEach((shape) => {
      const current = shapeMarginLevels.get(shape.id) ?? 0;
      const next = Math.max(0, current + delta);
      shapeMarginLevels.set(shape.id, next);
      setUniformMargins(shape, next);
    });

    await context.sync();
    showStatus(
      delta > 0
        ? `Text margins increased to ${shapeMarginLevels.get(shapes.items[0].id)}pt.`
        : `Text margins decreased to ${shapeMarginLevels.get(shapes.items[0].id)}pt.`
    );
  }).catch((err) => {
    showStatus(`Margin error: ${err.message}`);
    console.error(err);
  });
}

Office.actions.associate("IncreaseTextMargins", () => adjustTextMargins(MARGIN_STEP_PT));
Office.actions.associate("DecreaseTextMargins", () => adjustTextMargins(-MARGIN_STEP_PT));

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
