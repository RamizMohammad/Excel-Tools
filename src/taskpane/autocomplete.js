/**
 * Autocomplete/suggestions system for code editor
 */

const SUGGESTIONS_DB = {
  "xl.write": {
    label: "xl.write(df, anchor='A1', table=True)",
    description: "Write DataFrame to Excel",
    snippet: "xl.write(${1:df}, anchor='${2:A1}'${3:, table=True})",
    params: [
      { name: "df", type: "DataFrame", description: "The DataFrame to write" },
      { name: "anchor", type: "str", description: "Cell anchor (e.g., 'A1', 'B5')" },
      { name: "table", type: "bool", description: "Format as Excel table (default: True)" },
    ]
  },
  "xl.read": {
    label: "xl.read(range)",
    description: "Read from Excel range",
    snippet: "xl.read('${1:A1:D10}')",
    params: [
      { name: "range", type: "str", description: "Excel range (e.g., 'A1:D10')" }
    ]
  },
  "import pandas": {
    label: "import pandas as pd",
    description: "Import pandas library",
    snippet: "import pandas as pd"
  },
  "import numpy": {
    label: "import numpy as np",
    description: "Import numpy library",
    snippet: "import numpy as np"
  },
  "pd.DataFrame": {
    label: "pd.DataFrame(data)",
    description: "Create a pandas DataFrame",
    snippet: "pd.DataFrame({\n  ${1:\"column\"}: [${2:values}]\n})"
  }
};

export function updateSuggestions(textarea, suggestionsState) {
  const text = textarea.value;
  const cursorPos = textarea.selectionStart;
  const lineStart = text.lastIndexOf("\n", cursorPos - 1) + 1;
  const lineText = text.substring(lineStart, cursorPos);
  
  // Extract the word being typed
  const wordMatch = lineText.match(/\b(\w+(?:\.\w+)*(?:\()?)\s*$/);
  const word = wordMatch ? wordMatch[1] : "";
  
  if (!word || word.length < 1) {
    hideSuggestions(suggestionsState);
    return;
  }
  
  // Find matching suggestions
  const matches = Object.entries(SUGGESTIONS_DB)
    .filter(([key]) => key.toLowerCase().includes(word.toLowerCase()))
    .map(([key, value]) => ({ ...value, key, weight: calculateWeight(key, word) }))
    .sort((a, b) => b.weight - a.weight)
    .slice(0, 5);
  
  if (matches.length === 0) {
    hideSuggestions(suggestionsState);
    return;
  }
  
  suggestionsState.items = matches;
  suggestionsState.selectedIndex = 0;
  suggestionsState.visible = true;
  suggestionsState.word = word;
  renderSuggestions(suggestionsState);
}

function calculateWeight(key, typed) {
  const lowerKey = key.toLowerCase();
  const lowerTyped = typed.toLowerCase();
  
  if (lowerKey === lowerTyped) return 1000;
  if (lowerKey.startsWith(lowerTyped)) return 100;
  if (lowerKey.includes(lowerTyped)) return 10;
  return 1;
}

export function renderSuggestions(suggestionsState) {
  const panel = suggestionsState.panel;
  panel.innerHTML = "";
  
  suggestionsState.items.forEach((item, index) => {
    const div = document.createElement("div");
    div.className = "suggestion-item" + (index === suggestionsState.selectedIndex ? " selected" : "");
    
    const label = document.createElement("div");
    label.className = "suggestion-label";
    label.textContent = item.label;
    
    const desc = document.createElement("div");
    desc.className = "suggestion-desc";
    desc.textContent = item.description;
    
    div.appendChild(label);
    div.appendChild(desc);
    
    div.addEventListener("click", () => {
      suggestionsState.selectedIndex = index;
      selectSuggestion(suggestionsState.textarea, suggestionsState);
    });
    
    panel.appendChild(div);
  });
  
  panel.classList.add("visible");
}

export function selectSuggestion(textarea, suggestionsState) {
  const item = suggestionsState.items[suggestionsState.selectedIndex];
  if (!item) return;
  
  const text = textarea.value;
  const cursorPos = textarea.selectionStart;
  const lineStart = text.lastIndexOf("\n", cursorPos - 1) + 1;
  const lineText = text.substring(lineStart, cursorPos);
  
  // Remove the word being typed
  const wordMatch = lineText.match(/(\b\w+(?:\.\w+)*)\s*$/);
  const wordLen = wordMatch ? wordMatch[1].length : 0;
  
  const beforeCursor = text.substring(0, cursorPos - wordLen);
  const afterCursor = text.substring(cursorPos);
  
  textarea.value = beforeCursor + item.key + afterCursor;
  textarea.selectionStart = textarea.selectionEnd = beforeCursor.length + item.key.length;
  
  hideSuggestions(suggestionsState);
  
  // Trigger input event for updating cell code
  textarea.dispatchEvent(new Event("input", { bubbles: true }));
}

export function hideSuggestions(suggestionsState) {
  suggestionsState.visible = false;
  suggestionsState.panel.classList.remove("visible");
  suggestionsState.panel.innerHTML = "";
}
