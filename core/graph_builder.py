"""
Build and maintain the knowledge graph from scanned Excel nodes.
Supports both explicit (Meta.links) and implicit (cell-value match) edges.
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

import networkx as nx

from .scanner import node_id_from_path, parse_excel, scan_directory, scan_implicit_links


class GraphBuilder:
    def __init__(self, base_dir: Path) -> None:
        self.base_dir = base_dir
        self._nodes: dict[str, dict[str, Any]] = {}   # id -> node data
        self._graph: nx.DiGraph = nx.DiGraph()

    # ------------------------------------------------------------------
    # Build
    # ------------------------------------------------------------------

    def build(self, include_implicit: bool = True) -> None:
        """Full rebuild from scratch."""
        nodes = scan_directory(self.base_dir)
        self._nodes = {n["id"]: n for n in nodes}
        self._graph = nx.DiGraph()

        # Build lookup: title OR stem -> node_id
        name_to_id: dict[str, str] = {}
        for node in nodes:
            name_to_id[node["title"]] = node["id"]
            name_to_id[Path(node["id"]).stem] = node["id"]

        # Add all nodes
        for node in nodes:
            self._graph.add_node(node["id"], **node)

        # Explicit edges from Meta.links
        for node in nodes:
            for ref in node["links"]:
                target_id = name_to_id.get(ref)
                if target_id and target_id != node["id"]:
                    self._graph.add_edge(
                        node["id"], target_id,
                        relation_type="explicit",
                        weight=1.0,
                    )

        # Implicit edges (files without Meta sheet only)
        if include_implicit:
            known_names = set(name_to_id.keys())
            implicit_count = 0
            for node in nodes:
                if not node["has_meta"]:
                    file_path = self.base_dir / node["id"]
                    refs = scan_implicit_links(file_path, self.base_dir, known_names)
                    for ref in refs:
                        target_id = name_to_id.get(ref)
                        if target_id and target_id != node["id"]:
                            if not self._graph.has_edge(node["id"], target_id):
                                self._graph.add_edge(
                                    node["id"], target_id,
                                    relation_type="implicit",
                                    weight=0.3,
                                )
                                implicit_count += 1
            print(f"[KG] Implicit edges detected: {implicit_count}")

    # ------------------------------------------------------------------
    # Incremental updates
    # ------------------------------------------------------------------

    def update_node(self, node_id: str) -> None:
        """Re-parse a single file and update graph in place."""
        file_path = self.base_dir / node_id
        if not file_path.exists():
            return

        # Remove old node and its edges, then re-add
        if self._graph.has_node(node_id):
            self._graph.remove_node(node_id)

        node = parse_excel(file_path, self.base_dir)
        self._nodes[node_id] = node
        self._graph.add_node(node_id, **node)

        # Rebuild lookup from current state
        name_to_id: dict[str, str] = {}
        for nid, data in self._graph.nodes(data=True):
            name_to_id[data.get("title", nid)] = nid
            name_to_id[Path(nid).stem] = nid

        # Re-add explicit edges for this node
        for ref in node["links"]:
            target_id = name_to_id.get(ref)
            if target_id and target_id != node_id:
                self._graph.add_edge(node_id, target_id, relation_type="explicit", weight=1.0)

        # Re-add implicit edges for this node (only if no Meta sheet)
        known_names = set(name_to_id.keys())
        if not node["has_meta"]:
            refs = scan_implicit_links(file_path, self.base_dir, known_names)
            for ref in refs:
                target_id = name_to_id.get(ref)
                if target_id and target_id != node_id:
                    if not self._graph.has_edge(node_id, target_id):
                        self._graph.add_edge(node_id, target_id, relation_type="implicit", weight=0.3)

        # Re-add edges from other nodes that point to this node (explicit + implicit)
        new_title = node["title"]
        new_stem = Path(node_id).stem
        for other_id, other_data in self._graph.nodes(data=True):
            if other_id == node_id:
                continue
            # Explicit links from other nodes
            for ref in other_data.get("links", []):
                if ref in (new_title, new_stem):
                    self._graph.add_edge(other_id, node_id, relation_type="explicit", weight=1.0)
            # Implicit links from other nodes (only nodes without Meta)
            if not other_data.get("has_meta"):
                other_path = self.base_dir / other_id
                if other_path.exists():
                    refs = scan_implicit_links(other_path, self.base_dir, {new_title, new_stem})
                    if new_title in refs or new_stem in refs:
                        if not self._graph.has_edge(other_id, node_id):
                            self._graph.add_edge(other_id, node_id, relation_type="implicit", weight=0.3)

    def remove_node(self, node_id: str) -> None:
        """Remove a node and all its edges."""
        if self._graph.has_node(node_id):
            self._graph.remove_node(node_id)
        self._nodes.pop(node_id, None)

    # ------------------------------------------------------------------
    # Query
    # ------------------------------------------------------------------

    def to_d3_format(self, include_implicit: bool = False) -> dict[str, Any]:
        """Return graph as D3-compatible dict {nodes, links}."""
        degrees = dict(self._graph.degree())

        nodes = []
        for nid, data in self._graph.nodes(data=True):
            nodes.append({
                "id": nid,
                "title": data.get("title", nid),
                "tags": data.get("tags", []),
                "description": data.get("description", ""),
                "degree": degrees.get(nid, 0),
                "has_meta": data.get("has_meta", False),
            })

        links = []
        for source, target, data in self._graph.edges(data=True):
            rel_type = data.get("relation_type", "explicit")
            if not include_implicit and rel_type == "implicit":
                continue
            links.append({
                "source": source,
                "target": target,
                "type": rel_type,
                "weight": data.get("weight", 1.0),
            })

        return {"nodes": nodes, "links": links}

    def get_all_tags(self) -> list[str]:
        tags: set[str] = set()
        for _, data in self._graph.nodes(data=True):
            tags.update(data.get("tags", []))
        return sorted(tags)

    def get_node_data(self, node_id: str) -> dict[str, Any] | None:
        return self._nodes.get(node_id)

    @property
    def node_count(self) -> int:
        return len(self._nodes)
