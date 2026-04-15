let selectedNode = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.OneNote) {
        resetMap();
    }
});

function showError(msg) {
    const err = document.getElementById("error");
    err.innerHTML = "<strong>❌ " + msg + "</strong>";
    err.style.display = "block";
}

function selectNode(node) {
    document.querySelectorAll('.node').forEach(n => n.classList.remove('selected'));
    node.classList.add('selected');
    selectedNode = node;
}

function addChild() {
    const text = document.getElementById("nodeText").value.trim();
    if (!text) {
        return showError("Please type text in the box first");
    }

    const mindmap = document.getElementById("mindmap");
    let targetLevel = mindmap.lastElementChild;

    // If no levels yet or last level is empty, create new level
    if (!targetLevel || targetLevel.classList.contains('level') === false) {
        targetLevel = document.createElement("div");
        targetLevel.className = "level";
        mindmap.appendChild(targetLevel);
    }

    const newNode = document.createElement("div");
    newNode.className = "node";
    newNode.textContent = text;
    newNode.onclick = () => selectNode(newNode);

    targetLevel.appendChild(newNode);
    selectNode(newNode);

    document.getElementById("nodeText").value = "";
}

async function insertToOneNote() {
    const mindmap = document.getElementById("mindmap");
    if (!mindmap.children.length) return showError("Nothing to insert yet");

    try {
        document.getElementById("error").innerHTML = "<strong>Generating image... please wait</strong>";
        document.getElementById("error").style.display = "block";

        const canvas = await html2canvas(mindmap, {
            scale: 2.2,
            backgroundColor: "#ffffff",
            logging: false
        });

        const dataUrl = canvas.toDataURL("image/png");

        await OneNote.run(async (context) => {
            const page = context.application.getActivePage();
            const html = `
                <p><strong>Beautiful Mind Map</strong></p>
                <img src="${dataUrl}" style="max-width:100%; height:auto; border-radius:16px; box-shadow:0 15px 45px rgba(0,0,0,0.25);">
            `;
            page.addOutline(80, 120, html);
            await context.sync();
        });

        document.getElementById("error").style.display = "none";
        alert("✅ Mind map inserted successfully!");
    } catch (err) {
        showError("Insert failed: " + (err.message || "Unknown error"));
    }
}

function resetMap() {
    const mindmap = document.getElementById("mindmap");
    mindmap.innerHTML = '';

    // Create root level
    const rootLevel = document.createElement("div");
    rootLevel.className = "level";
    mindmap.appendChild(rootLevel);

    const root = document.createElement("div");
    root.className = "node root";
    root.textContent = "Central Idea";
    root.onclick = () => selectNode(root);

    rootLevel.appendChild(root);
    selectNode(root);

    document.getElementById("error").style.display = "none";
}