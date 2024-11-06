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

package nabu.data.excel;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TimeZone;

import javax.jws.WebParam;
import javax.jws.WebResult;
import javax.jws.WebService;
import javax.validation.constraints.NotNull;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import be.nabu.eai.api.NamingConvention;
import be.nabu.eai.repository.EAIResourceRepository;
import be.nabu.libs.property.api.Value;
import be.nabu.libs.services.api.ExecutionContext;
import be.nabu.libs.types.BaseTypeInstance;
import be.nabu.libs.types.ComplexContentWrapperFactory;
import be.nabu.libs.types.SimpleTypeWrapperFactory;
import be.nabu.libs.types.TypeUtils;
import be.nabu.libs.types.api.ComplexContent;
import be.nabu.libs.types.api.ComplexType;
import be.nabu.libs.types.api.DefinedType;
import be.nabu.libs.types.api.Element;
import be.nabu.libs.types.api.KeyValuePair;
import be.nabu.libs.types.api.SimpleType;
import be.nabu.libs.types.api.TypeInstance;
import be.nabu.libs.types.base.ComplexElementImpl;
import be.nabu.libs.types.base.SimpleElementImpl;
import be.nabu.libs.types.base.ValueImpl;
import be.nabu.libs.types.binding.excel.ExcelBinding;
import be.nabu.libs.types.java.BeanResolver;
import be.nabu.libs.types.properties.AliasProperty;
import be.nabu.libs.types.properties.LabelProperty;
import be.nabu.libs.types.properties.MaxOccursProperty;
import be.nabu.libs.types.structure.Structure;
import be.nabu.utils.excel.ExcelParser;
import be.nabu.utils.excel.ExcelUtils;
import be.nabu.utils.excel.FileType;
import be.nabu.utils.excel.MatrixUtils;
import be.nabu.utils.excel.Template;
import be.nabu.utils.excel.Template.Direction;
import be.nabu.utils.excel.ValueParserImpl;

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
	
	@SuppressWarnings("unchecked")
	@WebResult(name = "stream")
	public InputStream template(@WebParam(name = "stream") @NotNull InputStream stream, 
				@WebParam(name = "properties") Object variables, 
				@WebParam(name = "duplicateAll") Boolean duplicateAll, 
				@WebParam(name = "removeNonExistent") Boolean removeNonExistent,
				@WebParam(name = "direction") Direction direction, 
				@WebParam(name = "fileType") FileType fileType) throws IOException {
		
		if (fileType == null) {
			fileType = FileType.XLSX;
		}
		ExcelParser parser = new ExcelParser(stream, fileType, null);
		ByteArrayOutputStream target = new ByteArrayOutputStream();
		Map<String, Object> input = new HashMap<String, Object>();
		if (variables != null) {
			if (!(variables instanceof ComplexContent)) {
				variables = ComplexContentWrapperFactory.getInstance().getWrapper().wrap(variables);
			}
			if (variables != null) {
				for (Element<?> child : TypeUtils.getAllChildren(((ComplexContent) variables).getType())) {
					input.put(child.getName(), ((ComplexContent) variables).get(child.getName()));
				}
			}
		}
		Template.substitute(
			parser.getWorkbook(), 
			target, 
			input, 
			duplicateAll != null && duplicateAll, 
			direction == null ? Direction.VERTICAL : direction, 
			removeNonExistent != null && removeNonExistent
		);
		target.flush();
		return new ByteArrayInputStream(target.toByteArray());
	}
	
	@SuppressWarnings({ "unchecked", "rawtypes" })
	@WebResult(name = "unmarshalled")
	public Object unmarshal(@WebParam(name = "input") @NotNull InputStream input, @WebParam(name = "excelType") FileType type, @WebParam(name = "type") String typeId, @WebParam(name = "rotate") Boolean rotate, @WebParam(name = "includeEmptyResults") Boolean includeEmptyResults, @WebParam(name = "useHeaders") Boolean useHeaders, @WebParam(name = "validateHeaders") Boolean validateHeaders) throws IOException, ParseException {
		Workbook workbook = parse(input, type, null);
		ComplexType resolve;
		int headerRowIndex = 0;
		int contentRowIndex = headerRowIndex + 1;
		List<Integer> contentRowIndexes = new ArrayList<Integer>();
		if (typeId != null) {
			resolve = (ComplexType) EAIResourceRepository.getInstance().resolve(typeId);
		}
		else {
			ExcelParser excelParser = new ExcelParser(workbook);
			Structure dynamic = new Structure();
			FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
			for (String sheetName : excelParser.getSheetNames(false)) {
				Structure sheetType = new Structure();
				Sheet sheet = excelParser.getSheet(sheetName, false);
				if (sheet.getNumMergedRegions() > 0) {
					// check for merged headers in the beginning
					for (CellRangeAddress address : sheet.getMergedRegions()) {
						if (address.getFirstRow() <= headerRowIndex) {
							headerRowIndex = address.getLastRow() + 1;
						}
					}
				}
				contentRowIndex = headerRowIndex + 1;
				contentRowIndexes.add(contentRowIndex);
				// for naming
				Row headerRow = sheet.getRow(headerRowIndex);
				// for content type
				Row contentRow = sheet.getRow(contentRowIndex);
				for (int i = 0; i < headerRow.getLastCellNum(); i++) {
					Cell cell = headerRow.getCell(i);
					CellValue cellValue = evaluator.evaluate(cell);
					String cellName = cellValue.getStringValue();
					if (cellName == null || cellName.trim().isEmpty()) {
						break;
					}
					Class<?> simpleType = String.class;
					cell = contentRow.getCell(i);
					cellValue = evaluator.evaluate(cell);
					switch(cellValue.getCellType()) {
						case BOOLEAN:
							simpleType = Boolean.class;
						break;
						case NUMERIC:
							if (DateUtil.isCellInternalDateFormatted(cell) || DateUtil.isCellDateFormatted(cell)) {
								simpleType = Date.class;
							}
							else {
								simpleType = Double.class;
							}
						break;
					}
					sheetType.add(new SimpleElementImpl(NamingConvention.LOWER_CAMEL_CASE.apply(cellName), SimpleTypeWrapperFactory.getInstance().getWrapper().wrap(simpleType), sheetType));
				}
				dynamic.add(new ComplexElementImpl(NamingConvention.LOWER_CAMEL_CASE.apply(sheetName), sheetType, dynamic, new ValueImpl<String>(LabelProperty.getInstance(), sheetName), new ValueImpl<Integer>(MaxOccursProperty.getInstance(), 0)));
			}
			resolve = dynamic;
		}
		Iterator<Integer> iterator = contentRowIndexes.iterator();
		ComplexContent newInstance = resolve.newInstance();
		for (Element<?> child : TypeUtils.getAllChildren(resolve)) {
			Value<String> labelProperty = child.getProperty(LabelProperty.getInstance());
			// deprecated!
			Value<String> aliasProperty = child.getProperty(AliasProperty.getInstance());
			String sheetName = labelProperty == null ? (aliasProperty == null ? child.getName() : aliasProperty.getValue()) : labelProperty.getValue();
			if (sheetName.equals("*")) {
				sheetName = ".*";
			}
			// only use regex for sheet if we are not working with a dynamic type
			newInstance.set(child.getName(), toObject((ComplexType) child.getType(), workbook, sheetName, typeId != null, iterator.hasNext() ? iterator.next() : null, null, null, rotate, includeEmptyResults, useHeaders, validateHeaders, true));
		}
		return newInstance;
	}
	
	@SuppressWarnings("unchecked")
	@WebResult(name = "marshalled")
	public InputStream marshal(@WebParam(name = "data") Object data, @WebParam(name = "excelType") FileType type, @WebParam(name = "useHeaders") Boolean useHeaders, @WebParam(name = "timezone") TimeZone timezone) throws IOException, ParseException {
		if (data == null) {
			return null;
		}
		ExcelBinding binding = new ExcelBinding();
		if (useHeaders != null) {
			binding.setUseHeader(useHeaders);
		}
		ByteArrayOutputStream output = new ByteArrayOutputStream();
		if (!(data instanceof ComplexContent)) {
			data = ComplexContentWrapperFactory.getInstance().getWrapper().wrap(data);
			if (data == null) {
				throw new RuntimeException("Can not wrap data as complex content");
			}
		}
		binding.setTimezone(timezone);
		binding.marshal(output, (ComplexContent) data);
		return new ByteArrayInputStream(output.toByteArray());
	}
	
	@WebResult(name = "workbook")
	public Workbook newWorkbook(@WebParam(name = "fileType") FileType fileType) {
		return fileType == FileType.XLSX ? new XSSFWorkbook() : new HSSFWorkbook();
	}
	
	@WebResult(name = "stream")
	public InputStream toStream(@WebParam(name = "workbook") Workbook workbook) throws IOException {
		if (workbook == null) {
			return null;
		}
		ByteArrayOutputStream output = new ByteArrayOutputStream();
		workbook.write(output);
		return new ByteArrayInputStream(output.toByteArray());
	}
	
	@WebResult(name = "workbook")
	public Workbook toSheet(@WebParam(name = "workbook") Workbook workbook, @NotNull @WebParam(name = "sheet") String sheetName, @WebParam(name = "records") List<Object> objects, @WebParam(name = "useHeaders") Boolean useHeaders) {
		if (objects == null || objects.isEmpty()) {
			return workbook;
		}
		if (workbook == null) {
			workbook = newWorkbook(FileType.XLSX);
		}
		ExcelUtils.write(workbook, toMatrix(objects, useHeaders != null && useHeaders), sheetName, null, TimeZone.getDefault());
		return workbook;
	}
	
	@WebResult(name = "results")
	public List<Object> toObject(@NotNull @WebParam(name = "typeId") String typeId, @WebParam(name = "workbook") Workbook workbook, @NotNull @WebParam(name = "sheet") String sheetName, @WebParam(name = "useRegexForSheet") Boolean useRegex, @WebParam(name = "fromRow") Integer fromRow, @WebParam(name = "toRow") Integer toRow, @WebParam(name = "columnsToIgnore") List<Integer> columnsToIgnore, @WebParam(name = "rotate") Boolean rotate, @WebParam(name = "includeEmptyResults") Boolean includeEmptyResults, @WebParam(name = "useHeaders") Boolean useHeaders, @WebParam(name = "validateHeaders") Boolean validateHeaders, @WebParam(name = "trim") Boolean trim) throws IOException, ParseException {
		DefinedType resolved = executionContext.getServiceContext().getResolver(DefinedType.class).resolve(typeId);
		if (resolved == null) {
			throw new IllegalArgumentException("Could not resolve the complex type: " + typeId);
		}
		return toObject((ComplexType) resolved, workbook, sheetName, useRegex, fromRow, toRow, columnsToIgnore, rotate, includeEmptyResults, useHeaders, validateHeaders, trim);
	}
	
	@SuppressWarnings("unchecked")
	private static List<Object> toMatrix(List<Object> objects, boolean useHeaders) {
		List<Object> rows = new ArrayList<Object>();
		Collection<Element<?>> allChildren = null;
		if (useHeaders) {
			for (Object single : objects) {
				if (single == null) {
					continue;
				}
				else if (!(single instanceof ComplexContent)) {
					single = ComplexContentWrapperFactory.getInstance().getWrapper().wrap(single);
				}
				if (single == null) {
					throw new IllegalArgumentException("Could not wrap as complex content");
				}
				List<Object> row = new ArrayList<Object>();
				if (allChildren == null) {
					allChildren = TypeUtils.getAllChildren(((ComplexContent) single).getType());
				}
				for (Element<?> element : allChildren) {
					Value<String> label = element.getProperty(LabelProperty.getInstance());
					row.add(label != null ? label.getValue() : element.getName());
				}
				rows.add(row);
				break;
			}
		}
		for (Object single : objects) {
			if (single == null) {
				continue;
			}
			else if (!(single instanceof ComplexContent)) {
				single = ComplexContentWrapperFactory.getInstance().getWrapper().wrap(single);
			}
			if (single == null) {
				throw new IllegalArgumentException("Could not wrap as complex content");
			}
			List<Object> row = new ArrayList<Object>();
			if (allChildren == null) {
				allChildren = TypeUtils.getAllChildren(((ComplexContent) single).getType());
			}
			for (Element<?> element : allChildren) {
				row.add(((ComplexContent) single).get(element.getName()));
			}
			rows.add(row);
		}
		return rows;
	}
	
	private List<Object> toObject(ComplexType resolved, @WebParam(name = "workbook") Workbook workbook, @NotNull @WebParam(name = "sheet") String sheetName, @WebParam(name = "useRegexForSheet") Boolean useRegex, @WebParam(name = "fromRow") Integer fromRow, @WebParam(name = "toRow") Integer toRow, @WebParam(name = "columnsToIgnore") List<Integer> columnsToIgnore, @WebParam(name = "rotate") Boolean rotate, @WebParam(name = "includeEmptyResults") Boolean includeEmptyResults, @WebParam(name = "useHeaders") Boolean useHeaders, @WebParam(name = "validateHeaders") Boolean validateHeaders, @WebParam(name = "trim") Boolean trim) throws IOException, ParseException {
		// for backwards compatibility we trim by default
		if (trim == null) {
			trim = true;
		}
		List<Element<?>> children = new ArrayList<Element<?>>(TypeUtils.getAllChildren((ComplexType) resolved));
		// check if we have a key value pair list
		Iterator<Element<?>> iterator = children.iterator();
		TypeInstance keyValueInstance = new BaseTypeInstance(BeanResolver.getInstance().resolve(KeyValuePair.class));
		Element<?> keyValueElement = null;
		while (iterator.hasNext()) {
			Element<?> child = iterator.next();
			if (child.getType().isList(child.getProperties()) && TypeUtils.isSubset(new BaseTypeInstance(child.getType()), keyValueInstance)) {
				keyValueElement = child;
				// don't map normal values to it
				iterator.remove();
				break;
			}
		}
			
		ExcelParser excelParser = new ExcelParser(workbook);
		Sheet sheet = excelParser.getSheet(sheetName, useRegex == null ? false : useRegex);
		if (sheet == null) {
			throw new IllegalArgumentException("Can not find sheet: " + sheetName);
		}
		ValueParserImpl valueParser = new ValueParserImpl() {
			@Override
			public CellType getCellType(int cellIndex, Cell cell, CellValue value) {
				if (cellIndex < children.size()) {
					Element<?> element = children.get(cellIndex);
					if (element.getType() instanceof SimpleType) {
						// make sure we force strings to be parsed as string rather than a double
						Class<?> instanceClass = ((SimpleType<?>) element.getType()).getInstanceClass();
						if (String.class.isAssignableFrom(instanceClass)) {
							return CellType.STRING;
						}
					}
				}
				return super.getCellType(cellIndex, cell, value);
			}
		};
		// forcing a stringified version does not work, it will still format doubles with trailing ".0" or worse "1.5E7"
		// however, by forcing big decimals, they _do_ have a clean stringification path and should be compatible with normal doubles
		valueParser = new ValueParserImpl();
		valueParser.setUseBigDecimals(true);
		List<List<Object>> matrix = excelParser.matrix(sheet, valueParser);
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
		List<String> headers = new ArrayList<String>();
		for (int row = 0; row < matrix.size(); row++) {
			if (fromRow != null && row < fromRow) {
				continue;
			}
			else if (toRow != null && row >= toRow) {
				break;
			}
			if (useHeaders != null && useHeaders && ((fromRow == null && row == 0) || (fromRow != null && row == fromRow))) {
				int elementCounter = 0;
				for (int column = 0; column < matrix.get(row).size(); column++) {
					if (columnsToIgnore == null || !columnsToIgnore.contains(column)) {
						// add it to the headers for dynamic behavior
						headers.add(matrix.get(row).get(column) == null ? null : matrix.get(row).get(column).toString().trim());
						if (validateHeaders != null && validateHeaders) {
							// we ignore elements beyond the ones we can map, this could be empty cells or uninteresting data
							if (elementCounter >= children.size()) {
								continue;
							}
							Element<?> element = children.get(elementCounter++);
							Value<String> alias = element.getProperty(AliasProperty.getInstance());
							// don't check if it is an entire wildcard
							if (alias != null && alias.getValue().equals("*")) {
								continue;
							}
							String expectedName = alias == null ? element.getName() : alias.getValue();
							if (matrix.get(row).get(column) == null) {
								throw new ParseException("The actual header is null, expecting '" + expectedName + "'", row);
							}
							String actualName = matrix.get(row).get(column).toString().trim();
							if (!actualName.equals(expectedName)) {
								throw new ParseException("The actual header '" + actualName + "' does not match the expected '" + expectedName + "'", column);
							}
						}
					}
				}
				continue;
			}
			ComplexContent newInstance = ((ComplexType) resolved).newInstance();
			boolean isEmptyRow = true;
			int elementCounter = 0;
			List<ComplexContent> keyValuePairs = keyValueElement == null ? null : new ArrayList<ComplexContent>();
			for (int column = 0; column < matrix.get(row).size(); column++) {
				if (columnsToIgnore == null || !columnsToIgnore.contains(column)) {
					// we ignore elements beyond the ones we can map, this could be empty cells or uninteresting data
					if (elementCounter >= children.size()) {
						// if we have a key value element, add it there
						if (keyValueElement != null) {
							ComplexContent keyValuePair = ((ComplexType) keyValueElement.getType()).newInstance();
							keyValuePair.set("key", headers.get(elementCounter));
							Object value = matrix.get(row).get(column);
							if (value instanceof String) {
								// if we want to trim, we will save the trimmed version
								if (trim) {
									value = ((String) value).trim();
									if (((String) value).isEmpty()) {
										value = null;
									}
								}
								// otherwise we won't
								else if (((String) value).trim().isEmpty()) {
									value = null;
								}
							}
							keyValuePair.set("value", value);
							keyValuePairs.add(keyValuePair);
							elementCounter++;
							isEmptyRow = false;
							continue;
						}
						break;
					}
					Element<?> element = children.get(elementCounter++);
					try {
						Object value = matrix.get(row).get(column);
						if (value instanceof String) {
							// if we want to trim, we will save the trimmed version
							if (trim) {
								value = ((String) value).trim();
								if (((String) value).isEmpty()) {
									value = null;
								}
							}
							// otherwise we won't
							else if (((String) value).trim().isEmpty()) {
								value = null;
							}
						}
						newInstance.set(element.getName(), value);
					}
					catch (Exception e) {
						throw new IllegalArgumentException("Could not set field: " + element.getName(), e);
					}
					if (isEmptyRow && matrix.get(row).get(column) != null) {
						isEmptyRow = false;
					}
				}
			}
			if (keyValueElement != null) {
				newInstance.set(keyValueElement.getName(), keyValuePairs);
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
