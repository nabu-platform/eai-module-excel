package nabu.data.excel;

import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;

import javax.jws.WebParam;
import javax.jws.WebResult;
import javax.jws.WebService;
import javax.validation.constraints.NotNull;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import be.nabu.libs.services.api.ExecutionContext;
import be.nabu.libs.types.TypeUtils;
import be.nabu.libs.types.api.ComplexContent;
import be.nabu.libs.types.api.ComplexType;
import be.nabu.libs.types.api.DefinedType;
import be.nabu.libs.types.api.Element;
import be.nabu.utils.excel.ExcelParser;
import be.nabu.utils.excel.FileType;
import be.nabu.utils.excel.MatrixUtils;

@WebService
public class Services {
	
	private ExecutionContext executionContext;
	
	@SuppressWarnings("resource")
	@WebResult(name = "workbook")
	public Workbook parse(@NotNull @WebParam(name = "stream") InputStream input, @WebParam(name = "excelType") FileType type, @WebParam(name = "password") String password) throws IOException {
		return new ExcelParser(input, type, password).getWorkbook();
	}
	
	@SuppressWarnings("resource")
	@WebResult(name = "sheets")
	public List<String> sheets(@NotNull @WebParam(name = "workbook") Workbook workbook, @WebParam(name = "includeHidden") Boolean includeHidden) {
		return new ExcelParser(workbook).getSheetNames(includeHidden != null && includeHidden);
	}
	
	@SuppressWarnings("resource")
	@WebResult(name = "results")
	public List<Object> toObject(@NotNull @WebParam(name = "typeId") String typeId, @WebParam(name = "workbook") Workbook workbook, @NotNull @WebParam(name = "sheet") String sheetName, @WebParam(name = "useRegexForSheet") Boolean useRegex, @WebParam(name = "fromRow") Integer fromRow, @WebParam(name = "toRow") Integer toRow, @WebParam(name = "columnsToIgnore") List<Integer> columnsToIgnore, @WebParam(name = "rotate") Boolean rotate, @WebParam(name = "includeEmptyResults") Boolean includeEmptyResults) throws IOException, ParseException {
		DefinedType resolved = executionContext.getServiceContext().getResolver(DefinedType.class).resolve(typeId);
		if (resolved == null) {
			throw new IllegalArgumentException("Could not find the type: " + typeId);
		}
		if (!(resolved instanceof ComplexType)) {
			throw new IllegalArgumentException("The resolved type is not complex: " + typeId);
		}
		ExcelParser excelParser = new ExcelParser(workbook);
		Sheet sheet = excelParser.getSheet(sheetName, useRegex == null ? false : useRegex);
		if (sheet == null) {
			throw new IllegalArgumentException("Can not find sheet: " + sheetName);
		}
		List<List<Object>> matrix = excelParser.matrix(sheet);
		if (rotate != null && rotate) {
			matrix = MatrixUtils.rotate(matrix);
		}
		if (fromRow != null) {
			excelParser.setOffsetX(fromRow);
		}
		if (toRow != null) {
			excelParser.setMaxX(toRow);
		}
		List<Object> result = new ArrayList<Object>();
		List<Element<?>> children = new ArrayList<Element<?>>(TypeUtils.getAllChildren((ComplexType) resolved));
		for (int row = 0; row < matrix.size(); row++) {
			if (fromRow != null && row < fromRow) {
				continue;
			}
			else if (toRow != null && row >= toRow) {
				break;
			}
			ComplexContent newInstance = ((ComplexType) resolved).newInstance();
			boolean isEmptyRow = true;
			int elementCounter = 0;
			for (int column = 0; column < matrix.get(row).size(); column++) {
				if (columnsToIgnore == null || !columnsToIgnore.contains(column)) {
					// we ignore elements beyond the ones we can map, this could be empty cells or uninteresting data
					if (elementCounter >= children.size()) {
						break;
					}
					Element<?> element = children.get(elementCounter++);
					newInstance.set(element.getName(), matrix.get(row).get(column));
					if (isEmptyRow && matrix.get(row).get(column) != null) {
						isEmptyRow = false;
					}
				}
			}
			if (isEmptyRow && includeEmptyResults != null && includeEmptyResults) {
				result.add(null);
			}
			else if (!isEmptyRow) {
				result.add(newInstance);
			}
		}
		return result;
	}
}
