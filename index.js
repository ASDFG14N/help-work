const templateGarantia = "templates/garantia.docx";
const templateCredito = "templates/credito.docx";
const calculadoraTemplate = "templates/calculadoraTemplate.docx";

//Constantes
const dropdownButton = document.getElementById("dropdownDefaultButton");
const dropdownMenu = document.getElementById("dropdown");

const text = document.getElementById("text");
const dragArea = document.querySelector(".drag-area");

const now = new Date();

let selectedValue;
let contentUpload;

dropdownButton.addEventListener("click", function () {
  dropdownMenu.classList.toggle("hidden");
});

dropdownMenu.addEventListener("click", function (e) {
  if (e.target.tagName === "A") {
    const selectedOption = e.target.textContent.trim();
    selectedValue = e.target.getAttribute("data-value");

    dropdownButton.innerHTML = `${selectedOption}<svg class="w-2.5 h-2.5 ms-3" aria-hidden="true"
            xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 10 6">
            <path stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2"
              d="m1 1 4 4 4-4" />
          </svg>`;

    dropdownMenu.classList.add("hidden");
  }
});

//DragAndDrop
dragArea.addEventListener("dragover", (e) => {
  e.preventDefault();
});

dragArea.addEventListener("dragleave", (e) => {
  e.preventDefault();
});

dragArea.addEventListener("drop", (e) => {
  e.preventDefault();
  const file = e.dataTransfer.files[0];
  if (file != null) {
    try {
      text.innerHTML = `Archivo subido correctamente: <span class="font-semibold">${file.name}</span>`;
      const reader = new FileReader();

      reader.onload = function (e) {
        contentUpload = e.target.result;
      };

      reader.onerror = function (e) {
        text.textContent = `Error al procesar el archivo: ${error}`;
      };

      reader.readAsText(file);
    } catch (error) {
      text.textContent = `Error al procesar el archivo: ${error}`;
    }
  }
});

//envio del formulario
document.getElementById("myForm").addEventListener("submit", function (e) {
  e.preventDefault();
  switch (selectedValue) {
    case "0":
      read(contentUpload);
      setTimeout(() => {
        location.reload();
      }, 5000);
      break;
    case "1":
      const { tipoCambio } = e.target;
      readDataCalculator(contentUpload, tipoCambio.value);
      setTimeout(() => {
        location.reload();
      }, 5000);
      break;
    default:
      alert("Se debe seleccionar al menos una opción");
      break;
  }
});

const loadFile = (url, callback) => PizZipUtils.getBinaryContent(url, callback);

window.generate = function generate(inputFileName, outputFileName, data) {
  loadFile(inputFileName, function (error, content) {
    if (error) {
      throw error;
    }
    const zip = new PizZip(content);
    const doc = new window.docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });
    doc.render(data);

    const blob = doc.getZip().generate({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      compression: "DEFLATE",
    });
    saveAs(blob, outputFileName);
  });
};

const read = (data) => {
  const lines = data.split("\n");
  const output = lines.map((line) => line.split(";"));
  output.forEach((pair) => {
    if (pair[4] == 0) {
      generate(templateGarantia, pair[0] + "-garantia.docx", {
        operacion: pair[0],
        folio: pair[1],
        motor: pair[2],
        chasis: pair[3],
      });
    } else {
      generate(templateGarantia, pair[0] + "-garantia.docx", {
        operacion: pair[0],
        folio: pair[1],
        motor: pair[2],
        chasis: pair[3],
      });
      generate(templateCredito, pair[0] + "-credito.docx", {
        motor: pair[2],
        chasis: pair[3],
      });
    }
  });
};

//Calculadora registral
function roundToZero(number) {
  return parseFloat(
    (Math.floor(number * 10) / 10).toFixed(2).slice(0, -1) + "0"
  );
}

const calculatePayment = (originalValue, exchangeRate, isDollar) => {
  const dp = 12.3;
  const value = originalValue * exchangeRate;
  const di = roundToZero(value * (1.5 / 1000));
  const distr = di.toFixed(1) + "0";
  const vs = exchangeRate !== 1 ? value.toFixed(2) : null;
  const total = (dp + di).toFixed(1) + "0";

  if (isDollar) {
    return {
      dia: now.getDate(),
      mes: now.getMonth() + 1,
      anio: now.getFullYear(),
      hora: `${String(now.getHours()).padStart(2, "0")}:${String(
        now.getMinutes()
      ).padStart(2, "0")}`,
      tipo: exchangeRate,
      di: distr,
      v_original: originalValue,
      n_valor: vs,
      total,
    };
  } else {
    return {
      dia: now.getDate(),
      mes: now.getMonth() + 1,
      anio: now.getFullYear(),
      hora: `${String(now.getHours()).padStart(2, "0")}:${String(
        now.getMinutes()
      ).padStart(2, "0")}`,
      tipo: exchangeRate,
      di: distr,
      v_original: originalValue,
      n_valor: originalValue,
      total,
    };
  }
};

const readDataCalculator = (data, exchangeRate) => {
  const lines = data.split("\n");
  const output = lines.map((line) => {
    const pair = line.split(";");
    return [
      pair[0].trim(),
      pair[1].trim(),
      parseFloat(pair[2].replace("\r", "")),
    ];
  });
  output.forEach((pair) => {
    switch (pair[1]) {
      case "0":
        const objS = calculatePayment(pair[2], exchangeRate, false);
        generate(calculadoraTemplate, pair[0] + ".docx", objS);
        break;
      case "1":
        const objDo = calculatePayment(pair[2], exchangeRate, true);
        generate(calculadoraTemplate, pair[0] + ".docx", objDo);
        break;
      default:
        console.log("Algo salió mal");
        break;
    }
  });
};
