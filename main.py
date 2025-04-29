# feature_extractor_lightweight.py

import os
import ezdxf
import numpy as np
import cadquery as cq

class LightweightFeatureExtractor:

    def __init__(self):
        pass

    # --- STEP Parsing via cadquery ---
    def extract_step_features(self, filepath, filename_hint=None):
        part = cq.importers.importStep(filepath)
        bb = part.val().BoundingBox()

        features = {
            "width": bb.xlen,
            "height": bb.ylen,
            "depth": bb.zlen,
            "estimated_thickness": min(bb.xlen, bb.ylen, bb.zlen),
            "volume": part.val().Volume(),
            "area": part.val().Area(),
        }

        # Estimate bend count (very rough)
        faces = part.faces().vals()
        features["face_count"] = len(faces)
        features["bend_count"] = self._estimate_bends_by_face_normals(faces)

        # Estimate holes/slots based on faces with inner wires (very rough)
        features["hole_count"], features["slot_count"] = self._estimate_holes_and_slots(faces)

        material, finish = self._guess_material_and_finish(filename_hint)
        features["material"] = material
        features["surface_finish"] = finish

        return features

    def _estimate_bends_by_face_normals(self, faces):
        normals = [f.normalAt(0.5, 0.5) for f in faces]
        norm_vecs = [np.round((n.x, n.y, n.z), 2) for n in normals]
        unique_normals = set(norm_vecs)
        return max(0, len(unique_normals) - 1)

    def _estimate_holes_and_slots(self, faces):
        hole_count = 0
        slot_count = 0
        for face in faces:
            wires = face.Wires()
            if len(wires) > 1:  # inner wires = holes
                for i, wire in enumerate(wires):
                    if i == 0: continue  # skip outer wire
                    bbox = wire.val().BoundingBox()
                    ar = bbox.xlen / bbox.ylen if bbox.ylen != 0 else 0
                    if 0.2 < ar < 5:
                        slot_count += 1
                    else:
                        hole_count += 1
        return hole_count, slot_count

    # --- DXF Parsing ---
    def extract_dxf_features(self, filepath, filename_hint=None):
        doc = ezdxf.readfile(filepath)
        msp = doc.modelspace()

        total_length = 0
        hole_count = 0
        slot_count = 0

        for entity in msp:
            if entity.dxftype() == "LINE":
                p1 = np.array(entity.dxf.start)
                p2 = np.array(entity.dxf.end)
                total_length += np.linalg.norm(p2 - p1)
            elif entity.dxftype() == "CIRCLE":
                hole_count += 1
            elif entity.dxftype() == "LWPOLYLINE":
                if entity.is_closed and len(entity) > 4:
                    slot_count += 1

        material, finish = self._guess_material_and_finish(filename_hint)

        return {
            "total_cut_length": total_length,
            "hole_count": hole_count,
            "slot_count": slot_count,
            "material": material,
            "surface_finish": finish,
        }

    def _guess_material_and_finish(self, filename):
        if not filename:
            return "Unknown", "Raw"
        name = filename.lower()
        material = "Unknown"
        finish = "Raw"

        if "ss" in name or "stainless" in name:
            material = "Stainless Steel"
        elif "al" in name:
            material = "Aluminum"
        elif "ms" in name:
            material = "Mild Steel"

        if "painted" in name:
            finish = "Painted"
        elif "anodized" in name:
            finish = "Anodized"
        elif "galvanized" in name:
            finish = "Galvaniz"
def compare_features(self, f1: dict, f2: dict):
        comparison = {}
        keys = set(f1.keys()).union(f2.keys())
        for key in keys:
            if f1.get(key) != f2.get(key):
                comparison[key] = (f1.get(key), f2.get(key))
        return comparison

# Example use
if __name__ == "__main__":
    fx = LightweightFeatureExtractor()

    step_features = fx.extract_step_features("sample_part.step", filename_hint="bracket_SS_painted.step")
    print("STEP:", step_features)

    dxf_features = fx.extract_dxf_features("sample_part.dxf", filename_hint="bracket_SS_painted.dxf")
    print("DXF:", dxf_features)

    diff = fx.compare_features(step_features, dxf_features)
    print("DIFFERENCES:", diff)