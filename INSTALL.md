# Layout Shortcuts — install in 5 steps

Everything you need is in this folder. No git required.

## Step 1 — Create a new GitHub repo

1. Go to **https://github.com/new**
2. Repository name: **`layout-shortcuts`**
3. Set to **Public**
4. Do **not** add README, .gitignore, or license
5. Click **Create repository**

## Step 2 — Upload this folder

1. On the empty repo page, click **uploading an existing file**
2. Open Finder → **`~/Downloads/Layout-Shortcuts`**
3. Select **everything inside** (manifest.xml, shortcuts.json, .nojekyll, assets folder, src folder)
4. Drag it all into GitHub
5. Click **Commit changes**

## Step 3 — Turn on GitHub Pages

1. Go to **Settings → Pages**
2. **Source:** Deploy from a branch
3. **Branch:** `main` → folder **`/ (root)`**
4. Click **Save**
5. Wait 2 minutes, then open: **https://julievan.github.io/layout-shortcuts/shortcuts.json**  
   You should see `InsertTextBox` in the file.

> Important: use **Deploy from a branch**, not GitHub Actions.

## Step 4 — Install in PowerPoint

In Terminal:

```bash
mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
cp ~/Downloads/Layout-Shortcuts/manifest.xml ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
```

## Step 5 — Open PowerPoint

1. Quit PowerPoint completely (Cmd+Q)
2. Reopen and open a presentation
3. Go to **Home → Shortcuts** and click once
4. Test **Opt+Q** (add text box)

---

## All shortcuts

| Action | Shortcut |
|---|---|
| Align Left | Ctrl+Cmd+L |
| Align Right | Ctrl+Cmd+R |
| Align Center | Ctrl+Cmd+C |
| Align Top | Ctrl+Cmd+T |
| Align Bottom | Ctrl+Cmd+B |
| Align Middle | Ctrl+Cmd+M |
| Distribute Horizontally | Ctrl+Cmd+H |
| Distribute Vertically | Ctrl+Cmd+V |
| Same Size | Ctrl+Cmd+Z |
| Same Width | Ctrl+Cmd+E |
| Same Height | Cmd+Shift+E |
| Bring Forward | Opt+Cmd+Shift+F |
| Bring to Front | Cmd+Shift+F |
| Send Backward | Opt+Cmd+Shift+B |
| Send to Back | Cmd+Shift+B |
| Copy Position | Ctrl+1 |
| Paste Position | Ctrl+2 |
| Sticky Note | Ctrl+Cmd+0 |
| **Add Text Box** | **Opt+Q** |
| **Wrap Text in Shape** | **Opt+Ctrl+8** |
| **Paste Unformatted Text** | **Ctrl+Opt+T** |
| **Toggle Bullets** | **Opt+Ctrl+9** |

If PowerPoint asks which add-in owns a shortcut, choose **Layout Shortcuts**.
