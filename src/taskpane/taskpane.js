let nextId = 1;
let currentLevel = 0;

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

function createNode(text, isRoot = false, level = 0) {
    const mindmap = document.getElementById("mindmap");
    const node = document.createElement("div");
    node.className = "node" + (isRoot ? " root" : "");
    node.id = "node_" + nextId++;
    node.textContent = text;
    node.dataset.level = level;

    node.onclick = (e) => {
        e.stopPropagation();
        selectNode(node);
    };

    // Better positioning - horizontal tree style
    const leftPos = 80 + (level * 240);
    const topPos = 60 + (Math.random() * 80) + (level * 30);

    node.style.left = leftPos + "px";
    node.style.top = topPos + "px";

    mindmap.appendChild(node);
    return node;
}

function selectNode(node) {
    document.querySelectorAll('.node').forEach(n => n.classList.remove('selected'));
    node.classList.add('selected');
    window.selectedNode = node;
}

function addChild() {
    const text = document.getElementById("nodeText").value.trim();
    if (!text) return showError("Please type text in the box");

    let parent = window.selectedNode;
    if (!parent) {
        parent = document.getElementById("node_1"); // root
    }

    const parentLevel = parseInt(parent.dataset.level || 0);
    const newLevel = parentLevel + 1;

    const newNode = createNode(text, false, newLevel);
    selectNode(newNode);

    document.getElementById("nodeText").value = "";
}

async function insertToOneNote() {
    const mindmap = document.getElementById("mindmap");
    
    try {
        document.getElementById("error").innerHTML = "<strong>Generating image... (this may take a few seconds)</strong>";
        document.getElementById("error").style.display = "block";

        const canvas = await html2canvas(mindmap, {
            scale: 2,
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
        alert("✅ Beautiful mind map inserted successfully!");
        
    } catch (err) {
        showError("Insert failed: " + err.message);
    }
}

function resetMap() {
    const mindmap = document.getElementById("mindmap");
    mindmap.innerHTML = '';
    nextId = 1;

    // Create centered root node
    const root = createNode("Central Idea", true, 0);
    root.style.left = "calc(50% - 90px)";
    root.style.top = "60px";

    selectNode(root);
    document.getElementById("error").style.display = "none";
}