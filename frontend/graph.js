/* ============================================================
   graph.js — D3 force-directed knowledge graph
   ============================================================ */

// ── Tag colour palette ─────────────────────────────────────────
const TAG_COLORS = [
    '#7c3aed', '#2563eb', '#059669', '#d97706',
    '#dc2626', '#db2777', '#0891b2', '#65a30d',
    '#ea580c', '#0d9488',
];
const NO_TAG_COLOR = '#5a5a6e';

const tagColorMap = new Map();
let colorIndex = 0;

function getTagColor(tag) {
    if (!tagColorMap.has(tag)) {
        tagColorMap.set(tag, TAG_COLORS[colorIndex % TAG_COLORS.length]);
        colorIndex++;
    }
    return tagColorMap.get(tag);
}

// expose for panel.js
window.getTagColor = getTagColor;

// ── State ──────────────────────────────────────────────────────
let graphData = { nodes: [], links: [] };
let simulation = null;
let zoom = null;
let svgRoot = null;
let g = null;          // zoomable group
let selectedNodeId = null;
let localViewActive = false;
let localViewNeighbors = new Set();

// ── Init ───────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    initSVG();
    initToolbar();
    loadGraph();
    initWebSocket();
});

// ── SVG setup ──────────────────────────────────────────────────
function initSVG() {
    svgRoot = d3.select('#graph-svg');

    zoom = d3.zoom()
        .scaleExtent([0.05, 12])
        .on('zoom', (event) => g.attr('transform', event.transform));

    svgRoot.call(zoom)
        .on('dblclick.zoom', null) // disable default dblclick zoom
        .on('click', (event) => {
            // click on background → deselect / exit local view
            if (event.target.tagName === 'svg' || event.target.tagName === 'rect') {
                clearSelection();
            }
        });

    g = svgRoot.append('g').attr('id', 'scene');
    g.append('g').attr('class', 'links-layer');
    g.append('g').attr('class', 'nodes-layer');
    g.append('g').attr('class', 'labels-layer');

    simulation = d3.forceSimulation()
        .force('link', d3.forceLink().id(d => d.id).distance(d => d.type === 'implicit' ? 120 : 80).strength(0.5))
        .force('charge', d3.forceManyBody().strength(-220))
        .force('center', d3.forceCenter(0, 0))
        .force('collide', d3.forceCollide().radius(d => (d.radius || 8) + 12))
        .alphaDecay(0.028);
}

// ── Load graph from API ────────────────────────────────────────
async function loadGraph() {
    setLoading(true);
    try {
        const implicit = document.getElementById('implicit-toggle').checked;
        const res = await fetch(`/api/graph?include_implicit=${implicit}`);
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        graphData = await res.json();
        renderGraph();
        updateTagFilter();
        updateLegend();
        updateNodeCount();
    } catch (err) {
        console.error('Failed to load graph:', err);
    } finally {
        setLoading(false);
    }
}

// ── Render ────────────────────────────────────────────────────
function renderGraph() {
    const { nodes, links } = getFilteredData();

    if (nodes.length === 0) {
        document.getElementById('empty-state').classList.remove('hidden');
    } else {
        document.getElementById('empty-state').classList.add('hidden');
    }

    // Assign visual properties
    nodes.forEach(n => {
        n.color = n.tags.length > 0 ? getTagColor(n.tags[0]) : NO_TAG_COLOR;
        n.radius = Math.max(6, 4 + Math.sqrt(n.degree) * 4);
    });

    // ── Links ──
    const linkSel = g.select('.links-layer')
        .selectAll('line.link')
        .data(links, d => `${d.source.id ?? d.source}→${d.target.id ?? d.target}`)
        .join(
            enter => enter.append('line')
                .attr('class', d => `link ${d.type}`)
                .attr('marker-end', d => `url(#arrow-${d.type})`),
            update => update
                .attr('class', d => `link ${d.type}`)
                .attr('marker-end', d => `url(#arrow-${d.type})`),
            exit => exit.remove(),
        );

    // ── Nodes ──
    const nodeSel = g.select('.nodes-layer')
        .selectAll('circle.node')
        .data(nodes, d => d.id)
        .join(
            enter => enter.append('circle')
                .attr('class', 'node')
                .attr('r', d => d.radius)
                .attr('fill', d => d.color)
                .call(makeDrag())
                .on('click', onNodeClick)
                .on('dblclick', onNodeDblClick)
                .on('mouseover', onNodeHover)
                .on('mouseout', onNodeOut),
            update => update
                .attr('r', d => d.radius)
                .attr('fill', d => d.color),
            exit => exit.remove(),
        );

    // ── Labels ──
    g.select('.labels-layer')
        .selectAll('text.node-label')
        .data(nodes, d => d.id)
        .join(
            enter => enter.append('text')
                .attr('class', 'node-label')
                .text(d => truncate(d.title, 18)),
            update => update.text(d => truncate(d.title, 18)),
            exit => exit.remove(),
        );

    // Update simulation
    simulation.nodes(nodes);
    simulation.force('link').links(links);
    simulation.alpha(0.4).restart();

    simulation.on('tick', () => {
        g.select('.links-layer').selectAll('line.link')
            .attr('x1', d => d.source.x)
            .attr('y1', d => d.source.y)
            .attr('x2', d => clampLineEnd(d.source.x, d.target.x, d.target.radius || 8))
            .attr('y2', d => clampLineEnd(d.source.y, d.target.y, d.target.radius || 8));

        g.select('.nodes-layer').selectAll('circle.node')
            .attr('cx', d => d.x)
            .attr('cy', d => d.y);

        g.select('.labels-layer').selectAll('text.node-label')
            .attr('x', d => d.x)
            .attr('y', d => d.y + (d.radius || 8) + 13);
    });

    // Zoom to fit once simulation settles
    simulation.on('end', zoomToFit);
}

// Pull line end back so arrow doesn't overlap the circle
function clampLineEnd(s, t, r) {
    const d = t - s;
    const len = Math.sqrt((d) ** 2 + (d) ** 2) || 1;
    return t - (d / len) * r;
}

// ── Filtering ─────────────────────────────────────────────────
function getFilteredData() {
    const searchVal = document.getElementById('search-input').value.trim().toLowerCase();
    const tagVal = document.getElementById('tag-filter').value;

    let nodes = graphData.nodes;
    if (searchVal) {
        nodes = nodes.filter(n => n.title.toLowerCase().includes(searchVal));
    }
    if (tagVal) {
        nodes = nodes.filter(n => n.tags.includes(tagVal));
    }

    const nodeIds = new Set(nodes.map(n => n.id));
    const links = graphData.links.filter(l => {
        const s = l.source.id ?? l.source;
        const t = l.target.id ?? l.target;
        return nodeIds.has(s) && nodeIds.has(t);
    });

    return { nodes, links };
}

// ── Interactions ──────────────────────────────────────────────
function onNodeClick(event, d) {
    event.stopPropagation();
    selectedNodeId = d.id;
    applyHighlight(d.id);
    window.openPanel(d.id);
}

function onNodeDblClick(event, d) {
    event.stopPropagation();
    if (localViewActive && localViewNeighbors.has(d.id)) {
        clearSelection();
        return;
    }
    activateLocalView(d.id);
}

function onNodeHover(event, d) {
    showTooltip(event, d);
    if (!localViewActive) {
        g.select('.nodes-layer').selectAll('circle.node')
            .classed('dimmed', n => n.id !== d.id);
        g.select('.labels-layer').selectAll('text.node-label')
            .classed('dimmed', n => n.id !== d.id);
        g.select('.links-layer').selectAll('line.link')
            .classed('dimmed', l => {
                const s = l.source.id ?? l.source;
                const t = l.target.id ?? l.target;
                return s !== d.id && t !== d.id;
            })
            .classed('highlighted', l => {
                const s = l.source.id ?? l.source;
                const t = l.target.id ?? l.target;
                return s === d.id || t === d.id;
            });
    }
}

function onNodeOut() {
    hideTooltip();
    if (!localViewActive) {
        g.select('.nodes-layer').selectAll('circle.node').classed('dimmed', false);
        g.select('.labels-layer').selectAll('text.node-label').classed('dimmed', false);
        g.select('.links-layer').selectAll('line.link')
            .classed('dimmed', false).classed('highlighted', false);
    }
}

function activateLocalView(nodeId) {
    localViewActive = true;
    const neighbors = new Set([nodeId]);
    graphData.links.forEach(l => {
        const s = l.source.id ?? l.source;
        const t = l.target.id ?? l.target;
        if (s === nodeId) neighbors.add(t);
        if (t === nodeId) neighbors.add(s);
    });
    localViewNeighbors = neighbors;

    g.select('.nodes-layer').selectAll('circle.node')
        .classed('dimmed', d => !neighbors.has(d.id))
        .classed('highlighted', d => neighbors.has(d.id));
    g.select('.labels-layer').selectAll('text.node-label')
        .classed('dimmed', d => !neighbors.has(d.id))
        .classed('highlighted', d => neighbors.has(d.id));
    g.select('.links-layer').selectAll('line.link')
        .classed('dimmed', l => {
            const s = l.source.id ?? l.source;
            const t = l.target.id ?? l.target;
            return !neighbors.has(s) || !neighbors.has(t);
        });
}

function clearSelection() {
    selectedNodeId = null;
    localViewActive = false;
    localViewNeighbors.clear();
    g.select('.nodes-layer').selectAll('circle.node')
        .classed('dimmed', false).classed('selected', false).classed('highlighted', false);
    g.select('.labels-layer').selectAll('text.node-label')
        .classed('dimmed', false).classed('highlighted', false);
    g.select('.links-layer').selectAll('line.link')
        .classed('dimmed', false).classed('highlighted', false);
    window.closePanel();
}

function applyHighlight(nodeId) {
    g.select('.nodes-layer').selectAll('circle.node')
        .classed('selected', d => d.id === nodeId);
}

// Focus node in graph (called from panel link refs)
window.focusNodeByTitle = function(title) {
    const node = graphData.nodes.find(n => n.title === title || n.id.replace('.xlsx', '') === title);
    if (!node) return;
    onNodeClick({ stopPropagation: () => {} }, node);
    // Pan to it
    const w = svgRoot.node().clientWidth;
    const h = svgRoot.node().clientHeight;
    svgRoot.transition().duration(500).call(
        zoom.transform,
        d3.zoomIdentity.translate(w / 2 - node.x, h / 2 - node.y).scale(1.2)
    );
};

// ── Drag ──────────────────────────────────────────────────────
function makeDrag() {
    return d3.drag()
        .on('start', (event, d) => {
            if (!event.active) simulation.alphaTarget(0.3).restart();
            d.fx = d.x;
            d.fy = d.y;
        })
        .on('drag', (event, d) => {
            d.fx = event.x;
            d.fy = event.y;
        })
        .on('end', (event, d) => {
            if (!event.active) simulation.alphaTarget(0);
            d.fx = null;
            d.fy = null;
        });
}

// ── Zoom to fit ───────────────────────────────────────────────
function zoomToFit() {
    const svgEl = svgRoot.node();
    const bounds = g.node().getBBox();
    if (!bounds.width || !bounds.height) return;

    const w = svgEl.clientWidth || 800;
    const h = svgEl.clientHeight || 600;
    const scale = 0.8 / Math.max(bounds.width / w, bounds.height / h);
    const tx = w / 2 - scale * (bounds.x + bounds.width / 2);
    const ty = h / 2 - scale * (bounds.y + bounds.height / 2);

    svgRoot.transition().duration(600).call(
        zoom.transform,
        d3.zoomIdentity.translate(tx, ty).scale(scale)
    );
}

// ── Tooltip ───────────────────────────────────────────────────
const tooltipEl = document.getElementById('tooltip');

function showTooltip(event, d) {
    const tagHtml = d.tags.map(t =>
        `<span class="tooltip-tag" style="background:${getTagColor(t)}">${t}</span>`
    ).join('');

    tooltipEl.innerHTML = `
        <div class="tooltip-title">${d.title}</div>
        ${tagHtml ? `<div class="tooltip-tags">${tagHtml}</div>` : ''}
        <div class="tooltip-degree">連結數：${d.degree}</div>
    `;
    tooltipEl.classList.add('visible');
    moveTooltip(event);
}

function moveTooltip(event) {
    const x = event.clientX + 14;
    const y = event.clientY - 10;
    const tw = tooltipEl.offsetWidth;
    const th = tooltipEl.offsetHeight;
    const wx = window.innerWidth;
    const wy = window.innerHeight;
    tooltipEl.style.left = (x + tw > wx ? x - tw - 20 : x) + 'px';
    tooltipEl.style.top  = (y + th > wy ? y - th : y) + 'px';
}

document.addEventListener('mousemove', (e) => {
    if (tooltipEl.classList.contains('visible')) moveTooltip(e);
});

function hideTooltip() {
    tooltipEl.classList.remove('visible');
}

// ── Toolbar ───────────────────────────────────────────────────
function initToolbar() {
    document.getElementById('search-input').addEventListener('input', debounce(() => {
        renderGraph();
        updateNodeCount();
    }, 200));

    document.getElementById('tag-filter').addEventListener('change', () => {
        renderGraph();
        updateNodeCount();
    });

    document.getElementById('implicit-toggle').addEventListener('change', () => {
        loadGraph();
    });

    document.getElementById('refresh-btn').addEventListener('click', async () => {
        const btn = document.getElementById('refresh-btn');
        btn.textContent = '⟳ 掃描中…';
        btn.disabled = true;
        try {
            await fetch('/api/refresh', { method: 'POST' });
            await loadGraph();
        } finally {
            btn.textContent = '⟳ 重新整理';
            btn.disabled = false;
        }
    });

    document.getElementById('panel-close').addEventListener('click', clearSelection);

    const exportBtn = document.getElementById('export-btn');
    if (exportBtn) {
        exportBtn.addEventListener('click', async () => {
            exportBtn.textContent = '📥 產生中…';
            exportBtn.disabled = true;
            try {
                const res = await fetch('/api/export');
                if (!res.ok) throw new Error(`HTTP ${res.status}`);
                const blob = await res.blob();
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'knowledge-graph.html';
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            } catch (err) {
                alert('匯出失敗：' + err.message);
            } finally {
                exportBtn.textContent = '📥 匯出';
                exportBtn.disabled = false;
            }
        });
    }
}

function updateTagFilter() {
    const sel = document.getElementById('tag-filter');
    const currentVal = sel.value;

    // Rebuild options
    sel.innerHTML = '<option value="">全部標籤</option>';
    const allTags = new Set();
    graphData.nodes.forEach(n => n.tags.forEach(t => allTags.add(t)));
    [...allTags].sort().forEach(tag => {
        const opt = document.createElement('option');
        opt.value = tag;
        opt.textContent = tag;
        if (tag === currentVal) opt.selected = true;
        sel.appendChild(opt);
    });
}

function updateNodeCount() {
    const { nodes, links } = getFilteredData();
    document.getElementById('node-count').textContent =
        `${nodes.length} 個節點 · ${links.length} 條連結`;
}

function updateLegend() {
    const items = document.getElementById('legend-items');
    items.innerHTML = '';
    tagColorMap.forEach((color, tag) => {
        const div = document.createElement('div');
        div.className = 'legend-item';
        div.innerHTML = `<div class="legend-dot" style="background:${color}"></div><span>${tag}</span>`;
        items.appendChild(div);
    });
    if (!tagColorMap.size) {
        items.innerHTML = '<div style="color:#666;font-size:11px">（無標籤）</div>';
    }
}

// ── WebSocket ─────────────────────────────────────────────────
function initWebSocket() {
    const proto = location.protocol === 'https:' ? 'wss' : 'ws';
    const ws = new WebSocket(`${proto}://${location.host}/ws/updates`);

    ws.onmessage = (evt) => {
        const msg = JSON.parse(evt.data);
        if (msg.type === 'refresh' || msg.type === 'update') {
            loadGraph();
        }
    };

    ws.onclose = () => {
        setTimeout(initWebSocket, 3000); // auto-reconnect
    };
}

// ── Helpers ───────────────────────────────────────────────────
function truncate(str, len) {
    return str.length > len ? str.slice(0, len - 1) + '…' : str;
}

function setLoading(show) {
    document.getElementById('loading').classList.toggle('hidden', !show);
}

function debounce(fn, delay) {
    let t;
    return (...args) => {
        clearTimeout(t);
        t = setTimeout(() => fn(...args), delay);
    };
}
