Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Initialization code can go here
  }
});

/////////////////// Global Variables //////////////////

var apiKey;
var chatPromptValue = `Review this translation from Spanish to ensure it reads as if written by a native English speaker, maintaining the integrity of the original content as these are official translations. Make necessary corrections to enhance clarity and fluency, incorporating all changes directly into the original text without adding any extraneous information or explanationsso this is just a example how I wanted result the prompt text can be any thing 
        `;

var selectedText = ""; // Make sure this is updated properly
var result = "";

////////////////////////////////////////////////////////
///////////////////get api key//////////////////
async function fetchApiKey() {
  try {
    const response = await fetch("http://localhost:4000/api/openaikey/openaiKey");

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    apiKey = data.apiKey; // Set apiKey after fetching
    console.log("Fetched API Key:", apiKey);
  } catch (error) {
    console.error("Error fetching API key:", error);
  }
}

///////////////////////////////////////////////////

/////////prompt state///////////
// Define the initial value for the variable

const chatPromptInput = document.getElementById("chatPrompt");
const editCheckbox = document.getElementById("editCheckbox");

// Set initial input field value
chatPromptInput.value = chatPromptValue;

// Store initial value before editing
let initialValue = chatPromptValue;

// Listen for checkbox toggle to enable/disable editing
editCheckbox.addEventListener("change", function () {
  if (editCheckbox.checked) {
    // Enable editing and store initial value
    chatPromptInput.disabled = false;
    initialValue = chatPromptInput.value;
  } else {
    // Disable editing
    chatPromptInput.disabled = true;
  }
});

chatPromptInput.addEventListener("input", function () {
  chatPromptValue = chatPromptInput.value; // Update on each input change
  console.log("Updated chatPromptValue:", chatPromptValue); // Debugging log
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

// ChatGPT request logic
const sendChatGPTRequest = async () => {
  toggleLoading(true);

  await fetchApiKey();
  const requestBody = {
    model: "gpt-4-turbo",
    messages: [
      {
        role: "system",
        content: `You are a legal translation assistant. For each response, make sure to:
        - Preserve original legal terms and phrases (e.g., "Pursuant to the provisions of Article").
        - Maintain structure closely to the original example given.
        - Use formal legal language only.
        
        Here’s an example of the exact result style expected:

        "Pursuant to the provisions of Article 278, paragraph 3, subsection 2 of the CGP, in the current trial, there is no evidence to be presented, as the documentary evidence attached to the complaint has been considered with its appropriate legal probative value. Therefore, an ANTICIPATED JUDGMENT is issued in accordance with the aforementioned terms:"

        And this is the prompt text for above example result:

        "Pursuant to the provisions of Article 278, paragraph 3 N° 2° CGP, in the present trial there is no evidence to be presented, since the documentary evidence that was attached to the libel, was taken in its pertinent legal probative value, reason for which it issues ANTICIPATED JUDGMENT, with adhesion to what was stated, in the following terms:"

        so this is just a example how I wanted result the prompt text can be any thing 
        `,
      },
      {
        role: "user",
        content: `Review and translate the following text to ensure it aligns with the formal legal language as in the example provided. Make corrections directly in the text as necessary without altering key legal terms:

        Original Text: "${selectedText}".

        Additional instraction {${chatPromptValue}}
        
        
        `,
      },
    ],
    max_tokens: 1000,
  };

  const headers = {
    "Content-Type": "application/json",
    Authorization: `Bearer ${apiKey}`, // Replace with your API key
  };

  try {
    const response = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: headers,
      body: JSON.stringify(requestBody),
    });

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    result = data.choices[0].message.content;

    document.getElementById("chatGPTResponse").value = result;
    toggleLoading(false);
    // Replacing selected text after getting GPT result
    console.log("Full response content:", result);
  } catch (error) {
    console.error("Error fetching data from OpenAI:", error);
  }
};

// Function to call all necessary operations
async function callAllFunctions() {
  selectedText = await getSelectedText(); // Wait for the selected text
  if (selectedText) {
    await sendChatGPTRequest(); // Only send the request if selectedText is not null
  } else {
    console.log("No text selected, skipping ChatGPT request.");
  }
}
