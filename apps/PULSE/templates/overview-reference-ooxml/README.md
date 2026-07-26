# Overview export OOXML references

PowerPoint files are ZIP-based OOXML packages. These archives preserve the
complete supplied My Overview and Team Overview references, including slide
XML, relationships, theme, tables, and media.

The application does not copy the sample values from these files. Their visual
system is normalized in `assets/js/overview-pptx-template.js`, and the Overview
export fills that system with live PULSE data.

| Archive | Slides | Runtime use |
|---|---:|---|
| `my-overview-reference.zip` | 4 | Cover, personal metrics, assigned projects, task detail |
| `team-overview-reference.zip` | 9 | Cover, portfolio metrics, end items, inventory, project detail, workload |
