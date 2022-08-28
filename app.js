const express = require("express");
const app = express();
const port = 3000;
const path = require("path");
const rootDir = path.dirname(require.main.filename);
const IndexController = require("./IndexController");
const multerMiddleware = require("./middlewares/MulterMiddleware");

app.use(express.static(path.join(__dirname, "public")));
app.set("views engine", "ejs");

app.set("views", [path.join(rootDir, "views")]);

app.get("/", IndexController.getIndex);

app.post(
	"/readExcelFile",
	multerMiddleware("excel"),
	IndexController.readExcelData
);
app.get("/writeExcelFile", IndexController.writeExcelData);
app.get("/appendExcelFile", IndexController.appendToExistingExcelSheet);
app.get("/downloadRandomFile", IndexController.downloadFile);

app.listen(port, () => {
	console.log(
		`Excel Read Write app listening on port ${port}: http://localhost:${port}/`
	);
});
