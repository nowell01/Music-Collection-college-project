const DDLforChosen = document.getElementById("selectedOptions");
const DDLforAvail = document.getElementById("availOptions");

// Function to switch selected options from one dropdown to another
function switchOptions(event, senderDDL, receiverDDL) {
    event.preventDefault();

    const selectedOptions = Array.from(senderDDL.options).filter(option => option.selected);

    if (selectedOptions.length === 0) {
        alert("Nothing to move.");
        return;
    }

    selectedOptions.forEach(option => {
        senderDDL.remove(option.index);
        receiverDDL.appendChild(option);
    });
}

// Event handler functions
function addOptions(event) {
    switchOptions(event, DDLforAvail, DDLforChosen);
}

function removeOptions(event) {
    switchOptions(event, DDLforChosen, DDLforAvail);
}

// Assign event listeners
document.getElementById("btnLeft").addEventListener("click", addOptions);
document.getElementById("btnRight").addEventListener("click", removeOptions);

// Ensure all options in the chosen list are selected before submitting
document.getElementById("btnSubmit").addEventListener("click", () => {
    Array.from(DDLforChosen.options).forEach(option => {
        option.selected = true;
    });
});
