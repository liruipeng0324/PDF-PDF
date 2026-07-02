import argparse
import shutil
import sys

try:
    import pikepdf
    from pikepdf import OutlineItem
except ImportError:
    print("pikepdf is required. Install it with: python -m pip install pikepdf", file=sys.stderr)
    raise


def get_dict_value(dictionary, key):
    try:
        return dictionary.get(key)
    except Exception:
        return None


def destination_page_obj(destination):
    if destination is None:
        return None

    if isinstance(destination, pikepdf.Array):
        return destination[0] if len(destination) else None

    if isinstance(destination, pikepdf.Dictionary):
        if destination.get("/Type") == "/Page":
            return destination

        action_destination = get_dict_value(destination, "/D")
        if action_destination is not None:
            return destination_page_obj(action_destination)

    return None


def clone_pdf_value(value):
    if isinstance(value, pikepdf.Array):
        return pikepdf.Array([clone_pdf_value(item) for item in value])
    if isinstance(value, pikepdf.Dictionary):
        cloned = pikepdf.Dictionary()
        for key, item in value.items():
            cloned[key] = clone_pdf_value(item)
        return cloned
    return value


def clone_destination(destination, page_indexes, target_pdf):
    if destination is None:
        return None

    if isinstance(destination, pikepdf.Array):
        if len(destination) == 0:
            return None

        source_page = destination_page_obj(destination)
        if source_page is None:
            return None

        source_page_index = page_indexes.get(source_page.objgen)
        if source_page_index is None:
            return None

        cloned = pikepdf.Array([target_pdf.pages[source_page_index].obj])
        for index in range(1, len(destination)):
            cloned.append(clone_pdf_value(destination[index]))
        return cloned

    if isinstance(destination, pikepdf.Dictionary):
        if destination.get("/Type") == "/Page":
            source_page_index = page_indexes.get(destination.objgen)
            if source_page_index is None:
                return None
            return pikepdf.Array([target_pdf.pages[source_page_index].obj, pikepdf.Name("/Fit")])

        action_destination = get_dict_value(destination, "/D")
        if action_destination is not None:
            return clone_destination(action_destination, page_indexes, target_pdf)

    return None


def outline_destination(outline_node):
    direct_destination = get_dict_value(outline_node, "/Dest")
    if direct_destination is not None:
        return direct_destination

    action = get_dict_value(outline_node, "/A")
    if action is not None and get_dict_value(action, "/S") == "/GoTo":
        return get_dict_value(action, "/D")

    return None


def clone_outline_action(outline_node, page_indexes, target_pdf):
    action = get_dict_value(outline_node, "/A")
    if action is None or get_dict_value(action, "/S") != "/GoTo":
        return None

    destination = clone_destination(get_dict_value(action, "/D"), page_indexes, target_pdf)
    if destination is None:
        return None

    cloned_action = pikepdf.Dictionary()
    for key, value in action.items():
        if key == "/D":
            cloned_action[key] = destination
        else:
            cloned_action[key] = clone_pdf_value(value)
    return cloned_action


def clone_outline_node(outline_node, page_indexes, target_pdf):
    title = get_dict_value(outline_node, "/Title")
    if title is None:
        return None

    destination = clone_destination(outline_destination(outline_node), page_indexes, target_pdf)
    action = None
    if destination is None:
        action = clone_outline_action(outline_node, page_indexes, target_pdf)
    if destination is None and action is None:
        return None

    cloned = OutlineItem(str(title), destination=destination, action=action)

    child = get_dict_value(outline_node, "/First")
    while child is not None:
        cloned_child = clone_outline_node(child, page_indexes, target_pdf)
        if cloned_child is not None:
            cloned.children.append(cloned_child)
        child = get_dict_value(child, "/Next")

    return cloned


def source_outline_items(source_pdf, page_indexes, target_pdf):
    outlines = get_dict_value(source_pdf.Root, "/Outlines")
    if outlines is None:
        return []

    items = []
    current = get_dict_value(outlines, "/First")
    while current is not None:
        cloned = clone_outline_node(current, page_indexes, target_pdf)
        if cloned is not None:
            items.append(cloned)
        current = get_dict_value(current, "/Next")

    return items


def main():
    parser = argparse.ArgumentParser(description="Copy bookmarks from one PDF to another PDF with matching pages.")
    parser.add_argument("--source", required=True, help="Tagged PDF exported from Word.")
    parser.add_argument("--target", required=True, help="PDF that needs bookmarks.")
    parser.add_argument("--output", required=True, help="Final PDF path.")
    args = parser.parse_args()

    shutil.copyfile(args.target, args.output)

    with pikepdf.open(args.source) as source_pdf, pikepdf.open(args.output, allow_overwriting_input=True) as target_pdf:
        if len(source_pdf.pages) != len(target_pdf.pages):
            raise RuntimeError(
                f"Page count mismatch: source has {len(source_pdf.pages)} pages, target has {len(target_pdf.pages)} pages."
            )

        page_indexes = {page.objgen: index for index, page in enumerate(source_pdf.pages)}

        cloned_items = source_outline_items(source_pdf, page_indexes, target_pdf)

        with target_pdf.open_outline() as target_outline:
            target_outline.root.clear()
            for item in cloned_items:
                target_outline.root.append(item)

        target_pdf.save(args.output)


if __name__ == "__main__":
    main()
