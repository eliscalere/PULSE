# PULSE Diagram Refinement Plan

## Approach
Generate both diagrams as hand-crafted SVGs directly in code, then render to 4K PNG using a headless Chromium script.

## Diagram 1: Executive Architecture
- Light frosted-glass card system on white/near-white background with subtle grid texture
- Left → Center → Right flow: SharePoint Sources → PULSE Application Engine → Feature Modules → Outputs
- Engine as dominant center card with subtle blue border glow
- Connector arrows with action verb labels
- Metrics strip under title
- Callout card at bottom

## Diagram 2: Data Relationship Graph  
- Dark enterprise background (#0d1117)
- Force-directed layout manually optimized for 9 communities
- Colored convex hull regions per community
- Hub nodes enlarged with soft glow
- Edge thickness = relationship weight
- Metrics panel replacing old stats panel
- Community legend with colored swatches

## Communities from graphify data
1. Projects & Gantt (blue)
2. Meetings (teal)  
3. Documents (purple)
4. Travel (orange)
5. Notifications (yellow)
6. Users & Admin (green)
7. SharePoint Infrastructure (cyan)
8. Audit & Logging (red)
9. App Boot / System (gray)

## Export Strategy
- Write SVG files directly
- Use Node.js + Puppeteer for 4K PNG export (3840×2160)
