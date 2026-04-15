let selectedNode = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.OneNote) {
        console.log("✅ Add-in ready");
        selectedNode = document.getElementById("root");
    }
});

function showError(msg) {
    const err = document.getElementById("error");
    err.innerHTML = "<strong>❌ " + msg + "</strong>";
    err.style.display = "block";
}

function selectNode(el) {
    document.querySelectorAll('.node').forEach(n => n.style.borderColor = '#0078d4');
    el.style.borderColor = '#ff9800';
    selectedNode = el;
}

function addChild() {
    const text = document.getElementById("nodeText").value.trim();
    if (!text) return showError("Please enter node text");
    if (!selectedNode) return showError("Click on a node first");

    const mindmap = document.getElementById("mindmap");
    
    const newNode = document.createElement("div");
    newNode.className = "node";
    newNode.textContent = text;
    newNode.onclick = () => selectNode(newNode);
    
    // Simple visual branch (you can improve later with SVG lines)
    mindmap.appendChild(newNode);
    
    document.getElementById("nodeText").value = "";
    selectNode(newNode); // auto-select the new node
}

async function insertToOneNote() {
    const mindmap = document.getElementById("mindmap");
    
    try {
        showError("Generating beautiful image... please wait");

        const canvas = await html2canvas(mindmap, {
            scale: 2,           // higher resolution
            backgroundColor: "#ffffff",
            logging: false
        });

        const dataUrl = canvas.toDataURL("image/png");

        await OneNote.run(async (context) => {
            const page = context.application.getActivePage();
            const html = `
                <p><strong>Beautiful Mind Map</strong></p>
                <img src="${dataUrl}" style="max-width:100%; height:auto; border-radius:12px; box-shadow:0 10px 40px rgba(0,0,0,0.25);">
            `;
            page.addOutline(80, 120, html);
            await context.sync();
            
            document.getElementById("error").style.display = "none";
            alert("✅ Beautiful mind map inserted into OneNote!");
        });
    } catch (err) {
        showError("Insert failed: " + err.message);
    }
}

function resetMap() {
    if (confirm("Reset mind map?")) {
        document.getElementById("mindmap").innerHTML = '<div id="root" class="node" onclick="selectNode(this)">Central Idea</div>';
        document.getElementById("error").style.display = "none";
        selectedNode = document.getElementById("root");
    }
}