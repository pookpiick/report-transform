document.getElementById('file').addEventListener('change', function () {
  var name = document.getElementById('file-name');
  if (this.files.length) {
    name.textContent = this.files[0].name;
  } else {
    name.textContent = '';
  }
});

document.getElementById('previous').addEventListener('change', function () {
  var name = document.getElementById('previous-name');
  if (this.files.length) {
    name.textContent = this.files[0].name;
  } else {
    name.textContent = '';
  }
});
