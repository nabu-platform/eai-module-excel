/*
* Copyright (C) 2015 Alexander Verbruggen
*
* This program is free software: you can redistribute it and/or modify
* it under the terms of the GNU Lesser General Public License as published by
* the Free Software Foundation, either version 3 of the License, or
* (at your option) any later version.
*
* This program is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
* GNU Lesser General Public License for more details.
*
* You should have received a copy of the GNU Lesser General Public License
* along with this program. If not, see <https://www.gnu.org/licenses/>.
*/

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
