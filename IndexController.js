// Requiring the module
const xlxs = require("xlsx");
const fs = require("fs");

// Reading reader our test file

exports.getIndex = (req, res, next) => {
	res.render("index.ejs");
};

exports.downloadFile = (req, res, next) => {
	//res.download("Hello.txt");
};
exports.readExcelData = (req, res, next) => {
	const fileUpload = req.file;
	let fileLocation = fileUpload.destination + fileUpload.filename;
	// console.log(fileLocation);
	// return;
	let data = [];
	let file = xlxs.readFile("./" + fileLocation);
	const sheets = file.SheetNames;

	for (let i = 0; i < sheets.length; i++) {
		const temp = xlxs.utils.sheet_to_json(file.Sheets[file.SheetNames[i]]);
		temp.forEach(res => {
			data.push(res);
		});
	}

	// Printing data
	console.log(Object.values(data[0])); //To print all the values
	console.log(Object.keys(data[0])); // To print all the keys
};

exports.writeExcelData = async (req, res, next) => {
	// Requiring module

	// Reading our test file

	let student_data = [
		{
			Bank: "ICICI",
			Card: "Amazon Pay",
			Limit: "200000",
			DueDate: "15 th of every month",
		},
		{
			Bank: "HDFC",
			Card: "Moneyback+",
			Limit: "60000",
			DueDate: "23 th of every month",
		},
	];

	const ws = xlxs.utils.json_to_sheet(student_data);
	const wb = xlxs.utils.book_new();
	xlxs.utils.book_append_sheet(wb, ws, "Responses");
	let i = 0;
	let writeFile = true;
	try {
		while (writeFile) {
			if (fs.existsSync(`sampleData_${i}.xlsx`)) {
				i++;
				continue;
			} else {
				xlxs.writeFile(wb, `sampleData_${i}.xlsx`);
				writeFile = false;
			}
		}
	} catch (err) {
		console.error(err);
	}
	let downloadLink = `${__dirname}\\sampleData_${i}.xlsx`;
	const result = await res.download(downloadLink);
	if (result) {
		return res.redirect("/");
	}
};

exports.appendToExistingExcelSheet = async (req, res, next) => {
	// Requiring module

	// Reading our test file
	const file = xlxs.readFile("./Credit Cards Bill Payment Tracking.xlsx");
	const sheets = file.SheetNames;
	// Sample data set
	let student_data = [
		{
			Bank: "ICICI",
			Card: "Amazon Pay",
			Limit: "200000",
			DueDate: "15 th of every month",
		},
		{
			Bank: "HDFC",
			Card: "Moneyback+",
			Limit: "60000",
			DueDate: "23 th of every month",
		},
	];

	const ws = xlxs.utils.json_to_sheet(student_data);
	let i = sheets.length;
	xlxs.utils.book_append_sheet(file, ws, `Sheet${i + 1}`);

	// Writing to our file
	xlxs.writeFile(file, "./Credit Cards Bill Payment Tracking.xlsx");
	let downloadLink = `${__dirname}\\Credit Cards Bill Payment Tracking.xlsx`;
	const result = await res.download(downloadLink);
	if (result) {
		return res.redirect("/");
	}
};
