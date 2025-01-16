Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Initialization code can go here
  }
});

/////////////////// Global Variables //////////////////
var chatPromptValue = `Review this translation from Spanish to ensure it reads as if written by a native English speaker, maintaining the integrity of the original content as these are official translations. Make necessary corrections to enhance clarity and fluency, incorporating all changes directly into the original text without adding any extraneous information or explanationsso this is just a example how I wanted result the prompt text can be any thing 
        `;
var selectedText = ""; // Make sure this is updated properly
var result = "";

////////////////////////////////////////////////////////

/////////prompt state///////////
// Define the initial value for the variable

// Assign initial value to the input field
const chatPromptInput = document.getElementById("chatPrompt");
chatPromptInput.value = chatPromptValue;

// Checkbox to toggle edit mode
const editCheckbox = document.getElementById("editCheckbox");

// Listen for checkbox toggle to enable/disable editing
editCheckbox.addEventListener("change", function () {
  if (editCheckbox.checked) {
    chatPromptInput.disabled = false; // Enable editing
  } else {
    chatPromptInput.disabled = true; // Disable editing
    chatPromptValue = chatPromptInput.value; // Update the prompt value
    console.log("Final chatPromptValue:", chatPromptValue); // Log updated value
  }
});

//////////////////////////////////////////
////////////////loader//////////////////////////
function toggleLoading(show) {
  document.getElementById("loadingOverlay").style.display = show ? "flex" : "none";
}
/////////////////////////////////////////
// Function to get the selected text
function getSelectedText() {
  return Word.run((context) => {
    const selection = context.document.getSelection();
    selection.load("text");

    return context.sync().then(() => {
      const text = selection.text;
      console.log("Selected text:", text);

      if (text) {
        return text;
      } else {
        console.log("No text selected.");
        return null;
      }
    });
  }).catch((error) => {
    console.log("Error:", error);
    return null;
  });
}

// Function to replace the selected text with the result text
function replaceSelectedText() {
  return Word.run((context) => {
    const selection = context.document.getSelection();
    selection.load("text");

    return context.sync().then(() => {
      if (selection.text) {
        selection.insertText(result, Word.InsertLocation.replace);
        console.log("Replaced selected text with:", result);
      } else {
        console.log("No text selected, nothing to replace.");
      }
    });
  }).catch((error) => {
    console.log("Error:", error);
  });
}

///////////////////////////////////////////
async function fetchAnthropicMessage() {
  toggleLoading(true);
  try {
    const response = await fetch("https://copy-translation-backend.vercel.app/api/claude", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        content: `${chatPromptValue}

        "Pursuant to the provisions of Article 278, paragraph 3, subsection 2 of the CGP, in the current trial, there is no evidence to be presented, as the documentary evidence attached to the complaint has been considered with its appropriate legal probative value. Therefore, an ANTICIPATED JUDGMENT is issued in accordance with the aforementioned terms:"

        And this is the prompt text for above example result:

        "Pursuant to the provisions of Article 278, paragraph 3 N° 2° CGP, in the present trial there is no evidence to be presented, since the documentary evidence that was attached to the libel, was taken in its pertinent legal probative value, reason for which it issues ANTICIPATED JUDGMENT, with adhesion to what was stated, in the following terms:"

        {so this is just a example how I wanted result the prompt text can be any thing.}
        
        Review and translate the following text to ensure it aligns with the formal legal language as in the example provided. Make corrections directly in the text as necessary without altering key legal terms:

        Original Text: "${selectedText}". 
        final note: just give me translation of the selected text as requested in the prompt no need to add anything extra and also no need to add cotation around the translation`,
      }),
    });

    if (!response.ok) {
      throw new Error(`Server error: ${response.status} - ${response.statusText}`);
    }

    const data = await response.json();
    result = data.reply;

    document.getElementById("chatGPTResponse").value = result;
    toggleLoading(false);
    // Replacing selected text after getting GPT result
    console.log("Full response content:", result);
  } catch (error) {
    console.error("Error fetching message from backend:", error);
  }
}

///////////////////////////////
// Function to call all necessary operations
async function callAllFunctions() {
  selectedText = await getSelectedText(); // Wait for the selected text
  if (selectedText) {
    await fetchAnthropicMessage(); // Only send the request if selectedText is not null
  } else {
    console.log("No text selected, skipping ChatGPT request.");
  }
}
