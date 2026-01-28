import json
import xml.etree.ElementTree as ET
import os

# Paths
radek_json_path = "radek/cs.json"
current_json_path = "i18n/cs.json"
radek_xml_path = "radek/feeds2.xml"
output_feed_path = "i18n/feed_cs.txt"
output_json_path = "i18n/cs.json"

# 1. Merge JSONs
print("Merging JSONs...")
try:
    with open(radek_json_path, 'r', encoding='utf-8') as f:
        radek_data = json.load(f)
except Exception as e:
    print(f"Error reading radek json: {e}")
    radek_data = {}

try:
    with open(current_json_path, 'r', encoding='utf-8') as f:
        current_data = json.load(f)
except Exception as e:
    print(f"Error reading current json: {e}")
    current_data = {}

# Start with current data (to keep new keys), then update with Radek's data (priority)
# Wait, user said: "integra le traduzioni in radek dando priorit alle sue la guida prendi la sua e le traduzioni prendi le sue ma stringhe che mancano nelle sue prendi la mia traudzione"
# So: Base = Current (has new keys), Update with Radek (overwrites existing with his ver)
merged_data = current_data.copy()
merged_data.update(radek_data)

# Write merged JSON
with open(output_json_path, 'w', encoding='utf-8') as f:
    json.dump(merged_data, f, indent=2, ensure_ascii=False)
print(f"Merged JSON written to {output_json_path}")

# 2. Extract Feeds
print("Extracting Feeds...")
feeds = []
try:
    tree = ET.parse(radek_xml_path)
    root = tree.getroot()
    
    # Helper to process categories recursively if needed, but structure seems simple
    # <Category name="..."><Feed .../></Category> or <Category><Subcategory><Feed/></Subcategory></Category>
    
    def process_element(element, prefix=""):
        for child in element:
            if child.tag == 'Category' or child.tag.startswith('Subcategory'):
                name = child.get('name')
                new_prefix = f"{prefix}{name} > " if name else prefix
                process_element(child, new_prefix)
            elif child.tag == 'Feed':
                desc = child.get('Description')
                url = child.get('URL')
                if desc and url:
                    # Clean up title if it repeats prefix? No, user just wants list.
                    # Standard format: Title | URL
                    # Maybe include category in title?
                    # "NYT > Top Stories" example suggests Category > Title
                    full_title = f"{prefix}{desc}"
                    # Remove trailing " > " if desc is empty (unlikely)
                    if full_title.endswith(" > "):
                        full_title = full_title[:-3]
                    
                    feeds.append(f"{full_title} | {url}")

    process_element(root)

    with open(output_feed_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(feeds))
    print(f"Feeds written to {output_feed_path}")

except Exception as e:
    print(f"Error extracting feeds: {e}")
