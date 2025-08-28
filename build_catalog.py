import openpyxl

# Load your inventory Excel file
file_path = "cc.xlsx"   # update if your file name is different
wb = openpyxl.load_workbook(file_path)
sheet = wb.active

# Predefined categories
categories = {
    "Conduit and Fittings": [],
    "Fan and Appliances": [],
    "Hardware and Accessories": [],
    "Lighting": [],
    "Switches, Plates, and Accessories": [],
    "Switchgear and Protection": [],
    "Wires and Cables": []
}

# Read items from Excel (assuming Item in col A, Category in col B)
for row in sheet.iter_rows(min_row=2, values_only=True):
    item, category = row
    if item and category:
        if category in categories:
            categories[category].append(item)
        else:
            categories.setdefault(category, []).append(item)

# Build HTML
html_content = """
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Electrical Items Catalog</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 20px; }
    h1 { color: #2c3e50; }
    .category { margin-top: 20px; }
    .items { margin-left: 20px; display: none; }
    .search-box { margin-bottom: 20px; }
    input[type=text], input[type=email], textarea { padding: 8px; width: 300px; margin: 5px 0; }
    button { margin: 5px; padding: 8px 12px; cursor: pointer; }
    mark { background: yellow; font-weight: bold; }
    .selected { margin-top: 30px; padding: 10px; border: 1px solid #aaa; border-radius: 6px; }
    .selected-item { margin: 5px 0; padding: 5px; border: 1px solid #ddd; border-radius: 4px;
                     display: flex; justify-content: space-between; align-items: center; }
    .remove-btn { background: red; color: white; border: none; padding: 3px 8px; cursor: pointer; border-radius: 4px; }
    #success-message { display: none; margin-top: 20px; padding: 10px; border: 1px solid green;
                       color: green; border-radius: 6px; background: #eafbea; }
  </style>
  <script>
    function searchItems() {
      let input = document.getElementById('search').value.toLowerCase();
      let items = document.getElementsByClassName('item');
      for (let i = 0; i < items.length; i++) {
        let text = items[i].getAttribute("data-text").toLowerCase();
        if (input === "") {
          items[i].innerHTML = items[i].getAttribute("data-text");
          items[i].style.display = "";
        } else if (text.includes(input)) {
          let original = items[i].getAttribute("data-text");
          let regex = new RegExp("(" + input + ")", "ig");
          items[i].innerHTML = original.replace(regex, "<mark>$1</mark>");
          items[i].style.display = "";
        } else {
          items[i].style.display = "none";
        }
      }
    }

    function toggleAll(show) {
      let categories = document.getElementsByClassName('items');
      for (let i = 0; i < categories.length; i++) {
        categories[i].style.display = show ? "block" : "none";
      }
    }

    function addToSelected(itemText) {
      let selectedDiv = document.getElementById("selected-items");
      let existing = document.querySelector("#selected-items div[data-text='" + itemText + "']");
      if (existing) return;

      let div = document.createElement("div");
      div.className = "selected-item";
      div.setAttribute("data-text", itemText);

      let span = document.createElement("span");
      span.textContent = itemText;

      let btn = document.createElement("button");
      btn.textContent = "❌";
      btn.className = "remove-btn";
      btn.onclick = function() { div.remove(); updateHiddenField(); };

      div.appendChild(span);
      div.appendChild(btn);
      selectedDiv.appendChild(div);

      updateHiddenField();
    }

    function updateHiddenField() {
      let selected = [];
      let items = document.querySelectorAll("#selected-items div[data-text]");
      items.forEach(i => selected.push(i.getAttribute("data-text")));
      document.getElementById("selected-hidden").value = selected.join(", ");
    }

    function handleFormSubmit(event) {
      event.preventDefault();
      let form = event.target;

      fetch(form.action, {
        method: form.method,
        body: new FormData(form),
        headers: { 'Accept': 'application/json' }
      }).then(response => {
        if (response.ok) {
          document.getElementById("success-message").style.display = "block";
          form.reset();
          document.getElementById("selected-items").innerHTML = "";
          updateHiddenField();
        } else {
          alert("Oops! There was a problem submitting your form.");
        }
      }).catch(error => {
        alert("Error submitting form.");
      });
    }
  </script>
</head>
<body>
  <h1>Electrical Items Catalog</h1>
  <div class="search-box">
    <input type="text" id="search" onkeyup="searchItems()" placeholder="Search items...">
    <button onclick="toggleAll(true)">Expand All</button>
    <button onclick="toggleAll(false)">Collapse All</button>
  </div>

  <div id="catalog">
"""

# Add items by category
for category, items in categories.items():
    html_content += f"<div class='category'><h2 onclick=\"this.nextElementSibling.style.display=this.nextElementSibling.style.display==='none'?'block':'none'\">{category}</h2><div class='items'>"
    for item in items:
        html_content += f"<div class='item' data-text='{item}' onclick='addToSelected(\"{item}\")'>{item}</div>"
    html_content += "</div></div>"

# Add form + selected items
html_content += """
  </div>

  <div class="selected">
    <h2>Selected Items</h2>
    <div id="selected-items"></div>
  </div>

  <h2>Customer Details</h2>
  <form action="https://formspree.io/f/xnnbnkyb" method="POST" onsubmit="handleFormSubmit(event)">
    <label>Name:</label><br>
    <input type="text" name="name" required><br>
    <label>Email:</label><br>
    <input type="email" name="email" required><br>
    <label>Message:</label><br>
    <textarea name="message" rows="4" cols="50"></textarea><br>
    <input type="hidden" name="selected_items" id="selected-hidden">
    <button type="submit">Submit</button>
  </form>

  <div id="success-message">✅ Thank you! Your enquiry was sent successfully.</div>
</body>
</html>
"""

# Save file
with open("catalog.html", "w", encoding="utf-8") as f:
    f.write(html_content)


print("✅ Catalog website with inline success message generated: catalog.html")
