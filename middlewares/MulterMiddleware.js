const multer = require("multer");
const fileLocation = "public/file_uploads/";

const storageExcel = multer.diskStorage({
	destination: (req, file, cb) => {
		cb(null, fileLocation + "excel_sheet/");
	},
	filename: (req, file, cb) => {
		let splitted_name = file.originalname.split(".");
		let name1 = splitted_name[0].toLowerCase().split(" ").join("-");
		let name2 = splitted_name[1].toLowerCase(); //The file extension

		const full_file_name = name1 + Date.now() + "." + name2;
		cb(null, full_file_name);
	},
});

const multerMiddleware = storageOption => {
	// console.log(storageOption);
	if (storageOption == "excel") {
		return multer({ storage: storageExcel }).single("excel_file");
	}
};

module.exports = multerMiddleware;
