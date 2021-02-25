package be.nabu.eai.module.binding.excel;

import java.util.List;
import java.util.Map;

import be.nabu.eai.developer.api.TypeGenerator;
import be.nabu.eai.developer.api.TypeGeneratorTarget;
import be.nabu.libs.resources.api.ReadableResource;
import be.nabu.libs.types.api.ComplexContent;
import be.nabu.libs.types.structure.Structure;
import be.nabu.utils.excel.FileType;

public class ExcelTypeGenerator implements TypeGenerator {

	@Override
	public void requestUser(TypeGeneratorTarget target) {
		
	}

	@Override
	public boolean processResource(ReadableResource resource, TypeGeneratorTarget target) {
		if ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet".equals(resource.getContentType()) 
				|| (resource.getName() != null && resource.getName().toLowerCase().endsWith(".xlsx"))) {
			Map<Structure, List<ComplexContent>> parseExcel = GenerateExcelContextMenu.parseExcelAsContent(FileType.XLSX, resource, true);
			if (!parseExcel.isEmpty()) {
				target.generate(parseExcel);
			}
			return true;
		}
		else if (resource.getName() != null && resource.getName().toLowerCase().endsWith(".xls")) {
			Map<Structure, List<ComplexContent>> parseExcel = GenerateExcelContextMenu.parseExcelAsContent(FileType.XLS, resource, true);
			if (!parseExcel.isEmpty()) {
				target.generate(parseExcel);
			}
			return true;
		}
		return false;
	}

}
