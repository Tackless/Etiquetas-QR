function main(workbook: ExcelScript.Workbook) {
	let selectedSheet = workbook.getActiveWorksheet();
	
	// Bucle para eliminar cada imagen desde "Picture 1" hasta "Picture 1200"
	for (let i = 1; i <= 1200; i++) {
		let pictureName = `Picture ${i}`;
		let picture = selectedSheet.getShape(pictureName);
		
		// Verifica si la imagen existe antes de intentar eliminarla
		if (picture) {
			picture.delete();
		}
	}
}
