document.addEventListener("DOMContentLoaded", function () {
  // Get references to the form and result elements
  const form = document.getElementById("calculatorForm");
  const resultDiv = document.getElementById("result");

  // Function to handle the 'Calculate' button click
  async function calculate() {
    // Get the values of the input fields and parse them as floats
    const num1 = parseFloat(document.getElementById("num1").value);
    const num2 = parseFloat(document.getElementById("num2").value);

    // Check if the input values are valid numbers
    if (isNaN(num1) || isNaN(num2)) {
      alert("Please enter valid numbers.");
      return;
    }

    // Calculate the result and update the resultDiv
    const result = num1 + num2;
    resultDiv.innerText = `Result: ${result}`;

    // Send data to the server to write to Excel
    try {
      const response = await fetch("/calculate", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ num1, num2 }),
      });

      // Check if the request was successful
      if (!response.ok) {
        throw new Error("Failed to send data to the server.");
      } else {
        // Display a success alert
        alert("Numbers saved into Excel successfully!");
      }
    } catch (error) {
      console.error(error);
    }
  }

  // Function to handle the 'Print' button click
  function printToPDF() {
    // Send a request to the server to print the Excel sheet as PDF
    fetch("/print")
      .then((response) => response.json())
      .then((data) => {
        // Check if PDF generation was successful
        if (data.success) {
          alert("PDF generated successfully!");
        } else {
          alert("Failed to generate PDF.");
        }
      })
      .catch((error) => {
        console.error("Error while printing to PDF:", error);
      });
  }

  // Attach event listeners to buttons
  document
    .getElementById("calculateButton")
    .addEventListener("click", calculate);
  document.getElementById("printButton").addEventListener("click", printToPDF);
});
