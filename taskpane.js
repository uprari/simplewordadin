let fetchedData = '';

Office.onReady(() => {
  document.getElementById("fetchButton").onclick = fetchData;
  document.getElementById("displayButton").onclick = displayData;
  document.getElementById("writeButton").onclick = writeHtmlToWord;
});

async function fetchData() {
  const token = document.getElementById("token").value;
  const url = document.getElementById("url").value;

  if (!token || !url) {
    alert("Please provide both token and URL.");
    return;
  }

  try {
    const response = await fetch(url, {
      headers: {
        "Authorization": `Bearer ${token}`
      }
    });

    if (!response.ok) {
      throw new Error("Network response was not ok");
    }

    fetchedData = await response.text();
    alert("Data fetched successfully!");
  } catch (error) {
    console.error("Fetch error:", error);
    alert("Failed to fetch data.");
  }
}

function displayData() {
  if (!fetchedData) {
    alert("No data available. Please fetch data first.");
    return;
  }

  alert(`Fetched Data:\n\n${fetchedData.substring(0, 500)}...`);
}

function writeHtmlToWord() {
  if (!fetchedData) {
    alert("No data to write. Please fetch data first.");
    return;
  }

  Word.run(async (context) => {
    const body = context.document.body;
    body.insertHtml(fetchedData, Word.InsertLocation.end);
    await context.sync();
    alert("HTML content inserted into document.");
  }).catch(function (error) {
    console.error("Error: " + error);
  });
}

