let selectedNode = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.OneNote) {
        console.log("✅ Add-in ready");
        setTimeout(resetMap, 300);   // small delay for rendering
    }
});

function showError(msg) {
    const err = document.getElementById("error");
    err.textContent = msg;
    err.style.display = "block";
}

function showStatus(msg) {
    document.getElementById("status").textContent = msg;
}

function selectNode(node) {
    document.querySelectorAll('.node').forEach(n => n.classList.remove('selected'));
    node.classList.add('selected');
    selectedNode = node;
}

function addChild() {
    const text = document.getElementById("nodeText").value.trim();
    if (!text) return showError("Please type text in the box");

    if (!selectedNode) {
        selectedNode = document.querySelector('.node.root') || document.querySelector('.node');
    }

    const newNode = document.createElement("div");
    newNode.className = "node";
    newNode.textContent = text;
    newNode.onclick = () => selectNode(newNode);

    document.getElementById("mindmap").appendChild(newNode);
    selectNode(newNode);

    document.getElementById("nodeText").value = "";
    showStatus("Node added successfully");
}

async function insertToOneNote() {
    const mindmap = document.getElementById("mindmap");
    if (mindmap.children.length === 0) return showError("No mind map to insert yet");

    try {
        showStatus("Generating image... please wait");
        document.getElementById("error").style.display = "none";

        const canvas = await html2canvas(mindmap, {
            scale: 2.5,
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

        showStatus("✅ Inserted successfully!");
    } catch (err) {
        showError("Insert failed: " + err.message);
    }
}

function resetMap() {
    const mindmap = document.getElementById("mindmap");
    mindmap.innerHTML = '';

    const root = document.createElement("div");
    root.className = "node root";
    root.textContent = "Central Idea";
    root.onclick = () => selectNode(root);

    mindmap.appendChild(root);
    selectNode(root);

    document.getElementById("error").style.display = "none";
    showStatus("Click on a node, then type text and click 'Add Child'");
}