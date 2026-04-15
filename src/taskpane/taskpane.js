let nodes = [];
let nextId = 1;

Office.onReady((info) => {
    if (info.host === Office.HostType.OneNote) {
        resetMap();   // Create root node on load
    }
});

function showError(msg) {
    const err = document.getElementById("error");
    err.innerHTML = "<strong>❌ " + msg + "</strong>";
    err.style.display = "block";
}

function createNode(text, isRoot = false, parentId = null) {
    const mindmap = document.getElementById("mindmap");
    const node = document.createElement("div");
    node.className = "node" + (isRoot ? " root" : "");
    node.id = "node_" + nextId++;
    node.textContent = text;
    
    node.onclick = (e) => {
        e.stopPropagation();
        selectNode(node);
    };

    // Position nodes nicely
    if (isRoot) {
        node.style.left = "40%";
        node.style.top = "80px";
    } else {
        node.style.left = (Math.random() * 50 + 20) + "%";
        node.style.top = (Math.random() * 300 + 200) + "px";
    }

    mindmap.appendChild(node);
    nodes.push({ id: node.id, element: node, parentId });

    return node;
}

function selectNode(node) {
    document.querySelectorAll('.node').forEach(n => n.classList.remove('selected'));
    node.classList.add('selected');
    window.selectedNode = node;
}

function addChild() {
    const text = document.getElementById("nodeText").value.trim();
    if (!text) return showError("Please type something in the box");

    const parent = window.selectedNode || document.getElementById("node_1");
    if (!parent) return showError("No parent node found");

    const newNode = createNode(text, false, parent.id);
    selectNode(newNode);

    document.getElementById("nodeText").value = "";
}

async function insertToOneNote() {
    const mindmap = document.getElementById("mindmap");
    
    try {
        document.getElementById("error").innerHTML = "<strong>Generating beautiful image...</strong>";
        document.getElementById("error").style.display = "block";

        const canvas = await html2canvas(mindmap, {
            scale: 2.5,
            backgroundColor: "#ffffff",
            logging: false,
            useCORS: true
        });

        const dataUrl = canvas.toDataURL("image/png", 0.95);

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
        alert("✅ Beautiful mind map inserted into OneNote page!");
        
    } catch (err) {
        showError("Insert failed: " + err.message);
    }
}

function resetMap() {
    const mindmap = document.getElementById("mindmap");
    mindmap.innerHTML = '';
    nodes = [];
    nextId = 1;

    const root = createNode("Central Idea", true);
    selectNode(root);
    
    document.getElementById("error").style.display = "none";
}