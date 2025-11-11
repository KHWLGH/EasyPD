"""Power Delivery packet decoding utilities."""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple

from vendor_ids_dict import VENDOR_IDS

from device_comm import is_pdo, is_rdo


def is_sink_cap(pkg: Any) -> bool:
    """Return True if the packet looks like a Sink Capabilities message."""
    try:
        msg_type_value = pkg["Message Header"]["Message Type"].value()
        normalized = str(msg_type_value).strip().lower()
        return "sink" in normalized and "cap" in normalized
    except Exception:
        return False


def is_pdo_packet(pkg: Any) -> bool:
    if pkg is None:
        return False
    try:
        return bool(is_pdo(pkg)) and not is_sink_cap(pkg)
    except Exception:
        return False


def is_rdo_packet(pkg: Any) -> bool:
    if pkg is None:
        return False
    try:
        return bool(is_rdo(pkg))
    except Exception:
        return False


class PDParser:
    """PDO/RDO parser helpers."""

    @staticmethod
    def parse_pdo(pd_payload: Dict[str, Any]) -> List[Dict[str, str]]:
        pkg = pd_payload.get("data")
        if not pkg:
            return []

        entries: List[Dict[str, str]] = []
        try:
            is_extended = bool(pkg["Message Header"]["Extended"].value())
            data_objects = pkg[4].value() if is_extended else pkg[3].value()
        except Exception:
            return entries

        if not data_objects or not isinstance(data_objects, list):
            return entries

        for idx, obj in enumerate(data_objects):
            try:
                if getattr(obj, "value", lambda: None)() == "Empty PDO":
                    continue
            except Exception:
                pass

            summary = PDParser._safe_quick_call(obj, "quick_pdo")
            raw_bits = PDParser._safe_raw(obj)
            raw_display = PDParser._bits_to_hex(raw_bits)

            entries.append({
                "index": str(idx + 1),
                "summary": summary or "",
                "raw": raw_display,
            })

        return entries

    @staticmethod
    def parse_rdo(pd_payload: Dict[str, Any]) -> Dict[str, str]:
        pkg = pd_payload.get("data")
        info: Dict[str, str] = {}

        if pkg is None:
            info["summary"] = "Invalid RDO"
            return info

        try:
            data_objects = pkg[3].value()
        except Exception:
            info["summary"] = "Invalid RDO"
            return info

        if not data_objects or not isinstance(data_objects, list) or len(data_objects) == 0:
            info["summary"] = "Invalid RDO"
            return info

        target = data_objects[0]
        summary = PDParser._safe_quick_call(target, "quick_rdo")

        if not summary or summary == "Not a RDO":
            summary = "Invalid RDO"

        raw_bits = PDParser._safe_raw(target)
        raw_display = PDParser._bits_to_hex(raw_bits)

        try:
            position = target["Object Position"].value()
            position_text = f"Position: {position}"
        except Exception:
            position_text = None

        info["summary"] = summary
        if raw_display:
            info["raw"] = raw_display
        if position_text:
            info["details"] = position_text

        return info

    @staticmethod
    def _safe_raw(node: Any) -> str:
        try:
            return node.raw()
        except Exception:
            return ""

    @staticmethod
    def _bits_to_hex(bit_str: str) -> str:
        if not bit_str:
            return ""
        try:
            value = int(bit_str, 2)
            width = (len(bit_str) + 3) // 4
            return f"0x{value:0{width}X}"
        except Exception:
            return ""

    @staticmethod
    def _safe_quick_call(node: Any, attr: str) -> Optional[str]:
        try:
            method = getattr(node, attr)
            result = method()
            return result if isinstance(result, str) else str(result)
        except Exception:
            return None


class CableDataParser:
    """Structured VDM cable identity parser."""

    @staticmethod
    def parse(payload: Dict[str, Any]) -> Optional[List[Tuple[str, Any]]]:
        pkg = payload.get("data") if isinstance(payload, dict) else None
        if pkg is None:
            return None

        header = CableDataParser._get_metadata(pkg, "Message Header")
        if header is None:
            return None

        msg_type = CableDataParser._get_value(header, "Message Type")
        if msg_type != "Vendor_Defined":
            return None

        sop_value = CableDataParser._get_value(pkg, "SOP*")
        if sop_value not in ("SOP'", "SOP''"):
            return None

        data_objects = CableDataParser._get_metadata(pkg, "Data Objects")
        if data_objects is None:
            return None

        vdm_entries = data_objects.value()
        if not isinstance(vdm_entries, list) or not vdm_entries:
            return None

        vdm_header = vdm_entries[0]
        if not hasattr(vdm_header, "field") or vdm_header.field() != "VDM Header":
            return None

        vdm_type = CableDataParser._get_value(vdm_header, "VDM Type")
        if vdm_type != "Structured":
            return None

        command = CableDataParser._get_value(vdm_header, "Command")
        command_type = CableDataParser._get_value(vdm_header, "Command Type")
        if command != "Discover Identity" or command_type != "ACK":
            return None

        info: List[Tuple[str, Any]] = []

        svid = CableDataParser._get_value(vdm_header, "SVID")
        if svid:
            info.append(("cable_svid", str(svid)))

        info.append(("cable_source", str(sop_value)))
        info.append(("cable_command", f"{command} ({command_type})"))

        id_header = CableDataParser._find_by_field(vdm_entries, "ID Header VDO")
        if id_header is not None:
            vid = CableDataParser._get_value(id_header, "USB Vendor ID")
            vendor_name = CableDataParser._resolve_vendor_name(vid)
            if vid:
                vid_display = str(vid).upper()
                if vendor_name:
                    vid_display = f"{vid_display} ({vendor_name})"
                info.append(("cable_vendor_id", vid_display))

            cable_role = None
            if sop_value in ("SOP'", "SOP''"):
                cable_role = CableDataParser._get_value(id_header, "Product Type (Cable Plug/VPD)")
            if cable_role:
                info.append(("cable_role", str(cable_role)))

        product_vdo = CableDataParser._find_by_field(vdm_entries, "Product VDO")
        if product_vdo is not None:
            product_id = CableDataParser._get_value(product_vdo, "USB Product ID")
            bcd_device = CableDataParser._get_value(product_vdo, "bcdDevice")
            if product_id:
                info.append(("cable_product_id", str(product_id).upper()))
            if bcd_device:
                info.append(("cable_device_version", str(bcd_device).upper()))

        passive_vdo = CableDataParser._find_by_field(vdm_entries, "Passive Cable VDO")
        active_vdo_1 = CableDataParser._find_by_field(vdm_entries, "Active Cable VDO 1")
        active_vdo_2 = CableDataParser._find_by_field(vdm_entries, "Active Cable VDO 2")
        vpd_vdo = CableDataParser._find_by_field(vdm_entries, "VPD VDO")

        if passive_vdo is not None:
            info.append(("cable_type", "cable_type_passive"))
            info.extend(CableDataParser._extract_passive_info(passive_vdo))
        elif active_vdo_1 is not None:
            info.append(("cable_type", "cable_type_active"))
            info.extend(CableDataParser._extract_active_info(active_vdo_1, active_vdo_2))
        elif vpd_vdo is not None:
            info.append(("cable_type", "cable_type_vpd"))
            info.extend(CableDataParser._extract_vpd_info(vpd_vdo))

        return info if info else None

    @staticmethod
    def _get_metadata(node: Any, field: str):
        if node is None or not hasattr(node, "__getitem__"):
            return None
        try:
            child = node[field]
        except Exception:
            return None
        return child if hasattr(child, "value") else None

    @staticmethod
    def _get_value(node: Any, field: str):
        metadata_node = CableDataParser._get_metadata(node, field)
        if metadata_node is None:
            return None
        try:
            return metadata_node.value()
        except Exception:
            return None

    @staticmethod
    def _find_by_field(items: List[Any], field_name: str):
        for item in items:
            if hasattr(item, "field") and item.field() == field_name:
                return item
        return None

    @staticmethod
    def _resolve_vendor_name(vid_value: Any) -> Optional[str]:
        if vid_value is None:
            return None
        text = str(vid_value).strip()
        candidates: List[Any] = [text, text.upper()]

        upper_text = text.upper()
        if upper_text.startswith("0X"):
            candidates.append(upper_text[2:])
        else:
            candidates.append(f"0x{upper_text}")

        try:
            vid_int = int(text, 16)
        except (ValueError, TypeError):
            vid_int = None

        if vid_int is not None:
            candidates.extend([vid_int, f"{vid_int:04X}", f"0x{vid_int:04X}"])

        for candidate in candidates:
            if candidate in VENDOR_IDS:
                return VENDOR_IDS[candidate]
        return None

    @staticmethod
    def _extract_passive_info(node: Any) -> List[Tuple[str, Any]]:
        rows: List[Tuple[str, Any]] = []
        rows.append(("cable_connector", CableDataParser._fmt(CableDataParser._get_value(
            node, "USB Type-C plug to USB Type-C/Captive (Passive Cable)"))))
        rows.append(("cable_termination", CableDataParser._fmt(CableDataParser._get_value(
            node, "Cable Termination Type (Passive Cable)"))))
        rows.append(("cable_max_voltage", CableDataParser._fmt(CableDataParser._get_value(
            node, "Maximum VBUS Voltage (Passive Cable)"))))
        rows.append(("cable_current", CableDataParser._fmt(CableDataParser._get_value(
            node, "VBUS Current Handling Capability (Passive Cable)"))))
        rows.append(("cable_highest_speed", CableDataParser._fmt(CableDataParser._get_value(
            node, "USB Highest Speed (Passive Cable)"))))
        rows.append(("cable_latency", CableDataParser._fmt(CableDataParser._get_value(
            node, "Cable Latency (Passive Cable)"))))
        rows.append(("cable_supports_epr", CableDataParser._fmt(CableDataParser._get_value(
            node, "EPR Capable (Passive Cable)"))))
        return [row for row in rows if row[1]]

    @staticmethod
    def _extract_active_info(primary: Any, secondary: Any) -> List[Tuple[str, Any]]:
        rows: List[Tuple[str, Any]] = []
        rows.append(("cable_connector", CableDataParser._fmt(CableDataParser._get_value(
            primary, "USB Type-C plug to USB Type-C/Captive"))))
        rows.append(("cable_termination", CableDataParser._fmt(CableDataParser._get_value(
            primary, "Cable Termination Type (Active Cable)"))))
        rows.append(("cable_max_voltage", CableDataParser._fmt(CableDataParser._get_value(
            primary, "Maximum VBUS Voltage (Active Cable)"))))
        rows.append(("cable_current", CableDataParser._fmt(CableDataParser._get_value(
            primary, "VBUS Current Handling Capability (Active Cable)"))))
        rows.append(("cable_highest_speed", CableDataParser._fmt(CableDataParser._get_value(
            primary, "USB Highest Speed (Active Cable)"))))
        rows.append(("cable_supports_epr", CableDataParser._fmt(CableDataParser._get_value(
            primary, "EPR Capable (Active Cable)"))))

        sbu_supported = CableDataParser._get_value(primary, "SBU Supported")
        if sbu_supported is not None:
            rows.append(("cable_sbu", CableDataParser._fmt(sbu_supported)))

        if secondary is not None:
            rows.append(("cable_max_temp", CableDataParser._fmt(CableDataParser._get_value(
                secondary, "Maximum Operating Temperature"))))
            rows.append(("cable_shutdown_temp", CableDataParser._fmt(CableDataParser._get_value(
                secondary, "Shutdown Temperature"))))
            rows.append(("cable_usb4", CableDataParser._fmt(CableDataParser._get_value(
                secondary, "USB4 Supported"))))

        return [row for row in rows if row[1]]

    @staticmethod
    def _extract_vpd_info(node: Any) -> List[Tuple[str, Any]]:
        rows: List[Tuple[str, Any]] = []
        rows.append(("cable_max_voltage", CableDataParser._fmt(CableDataParser._get_value(
            node, "Maximum VBUS Voltage"))))
        rows.append(("cable_charge_through", CableDataParser._fmt(CableDataParser._get_value(
            node, "Charge Through Support"))))
        rows.append(("cable_charge_through_current", CableDataParser._fmt(CableDataParser._get_value(
            node, "Charge Through Current Support"))))
        rows.append(("cable_vbus_impedance", CableDataParser._fmt(CableDataParser._get_value(
            node, "VBUS Impedance"))))
        rows.append(("cable_ground_impedance", CableDataParser._fmt(CableDataParser._get_value(
            node, "Ground Impedance"))))
        return [row for row in rows if row[1]]

    @staticmethod
    def _fmt(value: Any) -> Any:
        if value is None:
            return ""
        if isinstance(value, bool):
            return "bool_true" if value else "bool_false"
        return str(value)


__all__ = [
    "is_sink_cap",
    "is_pdo_packet",
    "is_rdo_packet",
    "PDParser",
    "CableDataParser",
]
