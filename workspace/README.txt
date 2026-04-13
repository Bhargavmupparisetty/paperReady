PaperReady Workspace — Knowledge Base
=======================================

Place your reference documents here and PaperReady will automatically
index them for RAG-powered output generation.

SUPPORTED TEXT FORMATS
  .txt  .md  .rst  .csv  .log

SUPPORTED IMAGE FORMATS
  .png  .jpg  .jpeg  .bmp  .gif  .webp

IMAGE NAMING CONVENTION
  Name your images after the topic they represent so PaperReady
  can automatically embed them in presentations and documents.

  Examples:
    photosynthesis.png        → matched when topic contains "photosynthesis"
    neural_network.jpg        → matched for "neural network" topics
    climate_change_graph.png  → matched for "climate" topics

HOW IT WORKS
  1. You type: "create a powerpoint about photosynthesis"
  2. PaperReady scans this folder for any file matching "photosynthesis"
  3. It retrieves relevant text context from .txt/.md files
  4. It passes context + query to the local Phi-3 AI model
  5. It assembles your .pptx / .docx / .txt in the outputs/ folder

COMMANDS
  reload   — re-scan this folder without restarting
  exit     — quit PaperReady
