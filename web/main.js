var msgElement = document.getElementById("msg");

var loadingButtonUpload = document.getElementById("loading_upload");
var uploadBtn = document.getElementById("submit_upload_btn");

var folder_path_btn = document.getElementById("folder_path_btn");
var folder_path_val = document.getElementById("folder_path_val");
var folder_path_error = document.getElementById("folder_path_error");
var folder_path = "";

// to get the folder path of picture folder through python tkinter
folder_path_btn.addEventListener("click", async function () {
  folder_path = await eel.get_file_path()();
  folder_path_val.innerText = folder_path;
  folder_path_error.innerHTML = "";
});

function getCurrStatus(status) {
  msgElement.innerText += status + "\n";
  msgElement.style.color = "green";
}
eel.expose(getCurrStatus, "get_curr_status");

// uploading code
uploadBtn.addEventListener("click", function generateData(e) {
  e.preventDefault();
  msgElement.innerText = "";
  if (folder_path) {
    eel.start_driver_upload(folder_path)(viewMessage);
    loadingButtonUpload.style.display = "block";
    uploadBtn.style.display = "none";
  } else {
    folder_path_error.innerHTML = "Cannot be empty";
  }
});

function viewMessage(msg) {
  folder_path_val.innerText = "";
  msgElement.innerText = msg[0];
  msgElement.style.color = msg[1];
  loadingButtonUpload.style.display = "none";
  uploadBtn.style.display = "block";
}
