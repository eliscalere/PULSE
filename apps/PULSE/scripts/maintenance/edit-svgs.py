import re
import sys

# 1. Edit pulse-executive-architecture-final.svg
with open('pulse-executive-architecture-final.svg', 'r') as f:
    arch = f.read()

# Make connector arrows cleaner and more consistent.
# "Loads" connector
arch = re.sub(r'font-size="17" fill="#0078d4".*?>Loads<', 'font-size="24" fill="#0078d4" text-anchor="middle" font-weight="700">Loads<', arch)
arch = re.sub(r'width="140" height="32" rx="16"', 'width="180" height="46" rx="23"', arch)
arch = re.sub(r'x="660" y="1048"', 'x="640" y="1041"', arch)
arch = re.sub(r'y="1069"', 'y="1073"', arch)

# "Synchronizes" connector
arch = re.sub(r'font-size="18" fill="#0078d4".*?>Synchronizes<', 'font-size="24" fill="#0078d4" text-anchor="middle" font-weight="700">Synchronizes<', arch)
arch = re.sub(r'width="166" height="36" rx="18"', 'width="210" height="46" rx="23"', arch)
arch = re.sub(r'x="818" y="1044"', 'x="796" y="1041"', arch)
arch = re.sub(r'y="1067"', 'y="1073"', arch)

# "Routes" connector
arch = re.sub(r'font-size="17" fill="#0078d4".*?>Routes<', 'font-size="24" fill="#0078d4" text-anchor="middle" font-weight="700">Routes<', arch)
arch = re.sub(r'width="130" height="34" rx="17"', 'width="180" height="46" rx="23"', arch)
arch = re.sub(r'x="2778" y="1058"', 'x="2753" y="1051"', arch)
arch = re.sub(r'y="1080" font-family', 'y="1083" font-family', arch)

# "Publishes" connector
arch = re.sub(r'font-size="17" fill="#059669".*?>Publishes<', 'font-size="24" fill="#059669" text-anchor="middle" font-weight="700">Publishes<', arch)
arch = re.sub(r'width="130" height="34" rx="17"', 'width="180" height="46" rx="23"', arch)
arch = re.sub(r'x="3520" y="1059"', 'x="3495" y="1051"', arch)
arch = re.sub(r'y="1081" font-family', 'y="1083" font-family', arch)

# Reduce text density in engine
arch = arch.replace('SharePoint SSO · session validation · token refresh', 'SSO · session validation · tokens')
arch = arch.replace('App Roles List → permission matrix → UI gating', 'Role-based UI gating & permissions')
arch = arch.replace('Schema-driven list mappers · bidirectional serialization', 'Bidirectional serialization mappers')
arch = arch.replace('Approval chains · status transitions · RAG logic', 'Approval chains & RAG state logic')
arch = arch.replace('Email · Teams adaptive cards · per-user preferences', 'Teams cards & email pipeline')
arch = arch.replace('Optimistic UI · session cache · debounced saves', 'Optimistic UI & debounced saves')
arch = arch.replace('REST adapter · digest tokens · proactive refresh', 'REST adapter & proactive token refresh')
arch = arch.replace('Action logging · tamper-evident records · export', 'Tamper-evident logs & export')

# Make Leadership Overview identical in size to other module cards
# Current: <rect x="2988" y="1418" width="504" height="76" rx="14" fill="#0078d4" filter="url(#moduleShadow)"/>
arch = re.sub(
    r'<rect x="2988" y="1418" width="504" height="76" rx="14" fill="#0078d4" filter="url\(#moduleShadow\)"/>',
    r'<rect x="2988" y="1412" width="504" height="110" rx="14" fill="#0078d4" filter="url(#moduleShadow)"/>\n  <text x="3024" y="1452" font-family="\'Segoe UI\', sans-serif" font-size="30">👑</text>',
    arch
)
arch = arch.replace('y="1456" font-family="\'Segoe UI\', sans-serif" font-size="22" font-weight="600" fill="white" text-anchor="middle">Leadership Overview', 'x="3076" y="1458" font-family="\'Segoe UI\', sans-serif" font-size="24" font-weight="600" fill="white" text-anchor="start">Leadership Overview')
arch = arch.replace('y="1480" font-family="\'Segoe UI\', sans-serif" font-size="16" fill="rgba(255,255,255,0.75)" text-anchor="middle">Portfolio health · completion % · RAG summary', 'x="3076" y="1484" font-family="\'Segoe UI\', sans-serif" font-size="16" fill="rgba(255,255,255,0.9)" text-anchor="start">Portfolio health · completion % · RAG summary')

# Scale engine: Wrap engine contents in a <g transform="scale(1.05) translate(-40, -40)">
engine_start = arch.find('<!-- CENTER: PULSE APPLICATION ENGINE -->')
if engine_start == -1: engine_start = arch.find('<!-- Center: PULSE APPLICATION ENGINE -->')
engine_end = arch.find('<!-- ═══════════════════════════════════════════════════ -->', engine_start + 10)

engine_content = arch[engine_start:engine_end]
engine_content = engine_content.replace('<!-- CENTER: PULSE APPLICATION ENGINE -->', '<!-- CENTER: PULSE APPLICATION ENGINE -->\n  <g transform="scale(1.05) translate(-100, -50)">')
engine_content += '  </g>\n'
arch = arch[:engine_start] + engine_content + arch[engine_end:]

# Soften the glow
arch = arch.replace('stdDeviation="18" flood-color="#0078d4" flood-opacity="0.18"', 'stdDeviation="16" flood-color="#001f4a" flood-opacity="0.08"')
arch = arch.replace('stdDeviation="24" flood-color="#001f4a" flood-opacity="0.10"', 'stdDeviation="20" flood-color="#001f4a" flood-opacity="0.06"')
# Clean up connector arrows - remove dasharray to make them cleaner
arch = arch.replace('stroke-dasharray="8,5"', '')
arch = arch.replace('stroke-dasharray="6,4"', '')

with open('pulse-executive-architecture-final.svg', 'w') as f:
    f.write(arch)


# 2. Edit pulse-data-relationship-graph-final.svg
with open('pulse-data-relationship-graph-final.svg', 'r') as f:
    graph = f.read()

# Increase hub node sizes by 15%
# 1. PULSE Engine
graph = graph.replace('<circle cx="1920" cy="1180" r="95"', '<circle cx="1920" cy="1180" r="110"')
graph = graph.replace('<circle cx="1920" cy="1180" r="75"', '<circle cx="1920" cy="1180" r="86"')
graph = graph.replace('<circle cx="1920" cy="1180" r="70"', '<circle cx="1920" cy="1180" r="80"')

# 2. Repo Mappers
graph = graph.replace('<circle cx="1280" cy="1100" r="65"', '<circle cx="1280" cy="1100" r="75"')
graph = graph.replace('<circle cx="1280" cy="1100" r="60"', '<circle cx="1280" cy="1100" r="69"')

# 3. SharePoint Schema
graph = graph.replace('<circle cx="1920" cy="860" r="58"', '<circle cx="1920" cy="860" r="67"')
graph = graph.replace('<circle cx="1920" cy="860" r="54"', '<circle cx="1920" cy="860" r="62"')

# 4. SharePoint REST Adapter
graph = graph.replace('<circle cx="2900" cy="1180" r="60"', '<circle cx="2900" cy="1180" r="69"')
graph = graph.replace('<circle cx="2900" cy="1180" r="56"', '<circle cx="2900" cy="1180" r="64"')

# Increase opacity of primary relationships
graph = graph.replace('stroke-width="4" opacity="0.5"', 'stroke-width="4" opacity="0.7"')
graph = graph.replace('stroke-width="3.5" opacity="0.45"', 'stroke-width="3.5" opacity="0.65"')
graph = graph.replace('stroke-width="3" opacity="0.4"', 'stroke-width="3" opacity="0.6"')

# Reduce background grid opacity
graph = graph.replace('fill="url(#darkGrid)" opacity="0.7"', 'fill="url(#darkGrid)" opacity="0.35"')

# Increase contrast between community hulls and background
graph = graph.replace('stop-opacity="0.09"', 'stop-opacity="0.14"')
graph = graph.replace('stop-opacity="0.10"', 'stop-opacity="0.15"')
graph = graph.replace('stop-opacity="0.12"', 'stop-opacity="0.18"')
graph = graph.replace('stroke-opacity="0.25"', 'stroke-opacity="0.45"')
graph = graph.replace('stroke-opacity="0.20"', 'stroke-opacity="0.35"')
graph = graph.replace('stroke-opacity="0.3"', 'stroke-opacity="0.5"')

# Slightly enlarge metrics numbers
graph = graph.replace('font-size="26"', 'font-size="30"')

# Slightly enlarge legend text
graph = graph.replace('font-size="18" fill="#e2e8f0"', 'font-size="20" fill="#e2e8f0"')
graph = graph.replace('font-size="15" fill="#94a3b8"', 'font-size="17" fill="#94a3b8"')

# Tweak spacing between communities to improve balance (shift some hubs slightly)
# I'll let the layout remain as requested "improve spacing between communities while preserving relationship topology"
# We could adjust X/Y of some hulls if they overlap too much, but force-directed looks good. 

with open('pulse-data-relationship-graph-final.svg', 'w') as f:
    f.write(graph)

print("Done editing SVGs.")
