let nodes = [];
let nextId = 1;
let selectedNodeId = null;
let isDragging = false;
let dragNode = null;
let offsetX = 0, offsetY = 0;

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

function createNode(text, isRoot = false) {
    const mindmap = document.getElementById("mindmap");
    const node = document.createElement("div");
    node.className = "node" + (isRoot ? " root" : "");
    node.id = "node_" + nextId++;
    node.textContent = text;

    // Mouse drag support
    node.addEventListener("mousedown", startDrag);

    mindmap.appendChild(node);
    nodes.push({ id: node.id, element: node, children: [], text: text });

    return node;
}

function startDrag(e) {
    dragNode = e.currentTarget;
    isDragging = true;
    const rect = dragNode.getBoundingClientRect();
    const containerRect = document.getElementById("mindmap-container").getBoundingClientRect();
    
    offsetX = e.clientX - rect.left;
    offsetY = e.clientY - rect.top;

    document.addEventListener("mousemove", doDrag);
    document.addEventListener("mouseup", stopDrag);
}

function doDrag(e) {
    if (!isDragging || !dragNode) return;
    const container = document.getElementById("mindmap-container");
    const contRect = container.getBoundingClientRect();

    let newLeft = e.clientX - contRect.left - offsetX;
    let newTop = e.clientY - contRect.top - offsetY;

    // Prevent dragging too far out
    newLeft = Math.max(20, Math.min(newLeft, contRect.width - 200));
    newTop = Math.max(20, Math.min(newTop, contRect.height - 100));

    dragNode.style.left = newLeft + "px";
    dragNode.style.top = newTop + "px";

    drawLines();
}

function stopDrag() {
    isDragging = false;
    dragNode = null;
    document.removeEventListener("mousemove", doDrag);
    document.removeEventListener("mouseup", stopDrag);
}

function selectNode(id) {
    document.querySelectorAll('.node').forEach(n => n.classList.remove('selected'));
    const nodeEl = document.getElementById(id);
    if (nodeEl) nodeEl.classList.add('selected');
    selectedNodeId = id;
    drawLines();
}

function addChild() {
    const text = document.getElementById("nodeText").value.trim();
    if (!text) return showError("Please type node text");

    let parentId = selectedNodeId || "node_1";
    const parent = nodes.find(n => n.id === parentId);
    if (!parent) return showError("Parent not found");

    const newNodeEl = createNode(text);
    parent.children.push(newNodeEl.id);

    // Place to the right with spacing
    const parentEl = document.getElementById(parentId);
    const pRect = parentEl.getBoundingClientRect();
    const cRect = document.getElementById("mindmap-container").getBoundingClientRect();

    newNodeEl.style.left = (pRect.left - cRect.left + 240) + "px";
    newNodeEl.style.top = (pRect.top - cRect.top + (parent.children.length - 1) * 90) + "px";

    selectNode(newNodeEl.id);
    drawLines();

    document.getElementById("nodeText").value = "";
}

function drawLines() {
    const svg = document.getElementById("lines");
    svg.innerHTML = "";

    nodes.forEach(node => {
        if (node.children.length === 0) return;
        const parentEl = node.element;
        const pRect = parentEl.getBoundingClientRect();
        const contRect = document.getElementById("mindmap-container").getBoundingClientRect();

        node.children.forEach(childId => {
            const childEl = document.getElementById(childId);
            if (!childEl) return;
            const cRect = childEl.getBoundingClientRect();

            const x1 = pRect.left - contRect.left + pRect.width / 2;
            const y1 = pRect.top - contRect.top + pRect.height / 2;
            const x2 = cRect.left - contRect.left + 15;
            const y2 = cRect.top - contRect.top + cRect.height / 2;

            const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
            path.setAttribute("class", "connector");
            path.setAttribute("d", `M ${x1} ${y1} Q ${x1 + 80} ${y1}, ${x2} ${y2}`);
            svg.appendChild(path);
        });
    });
}

async function insertToOneNote() {
    const container = document.getElementById("mindmap-container");
    try {
        document.getElementById("error").innerHTML = "<strong>Generating full image... (may take a few seconds)</strong>";
        document.getElementById("error").style.display = "block";

        const canvas = await html2canvas(container, {
            scale: 2.5,
            backgroundColor: "#ffffff",
            logging: false,
            scrollX: 0,
            scrollY: 0,
            width: container.scrollWidth,
            height: container.scrollHeight
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
        alert("✅ Full mind map inserted into OneNote!");
    } catch (err) {
        showError("Insert failed: " + err.message);
    }
}

function resetMap() {
    const mindmap = document.getElementById("mindmap");
    const svg = document.getElementById("lines");
    mindmap.innerHTML = '';
    svg.innerHTML = '';
    nodes = [];
    nextId = 1;

    const root = createNode("Central Idea", true);
    root.style.left = "420px";   // centered
    root.style.top = "80px";

    selectNode(root.id);
    document.getElementById("error").style.display = "none";
}