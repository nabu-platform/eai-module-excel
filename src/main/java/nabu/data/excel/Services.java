package nabu.data.excel;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.jws.WebParam;
import javax.jws.WebResult;
import javax.jws.WebService;
import javax.validation.constraints.NotNull;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import be.nabu.eai.repository.EAIResourceRepository;
import be.nabu.libs.property.api.Value;
import be.nabu.libs.services.api.ExecutionContext;
import be.nabu.libs.types.ComplexContentWrapperFactory;
import be.nabu.libs.types.TypeUtils;
import be.nabu.libs.types.api.ComplexContent;
import be.nabu.libs.types.api.ComplexType;
import be.nabu.libs.types.api.DefinedType;
import be.nabu.libs.types.api.Element;
import be.nabu.libs.types.binding.excel.ExcelBinding;
import be.nabu.libs.types.properties.AliasProperty;
import be.nabu.utils.excel.ExcelParser;
import be.nabu.utils.excel.FileType;
import be.nabu.utils.excel.MatrixUtils;
import be.nabu.utils.excel.Template;
import be.nabu.utils.excel.Template.Direction;

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
	
	@WebResult(name = "unmarshalled")
	public Object unmarshal(@WebParam(name = "input") @NotNull InputStream input, @WebParam(name = "excelType") FileType type, @NotNull @WebParam(name = "type") String typeId, @WebParam(name = "rotate") Boolean rotate, @WebParam(name = "includeEmptyResults") Boolean includeEmptyResults, @WebParam(name = "useHeaders") Boolean useHeaders, @WebParam(name = "validateHeaders") Boolean validateHeaders) throws IOException, ParseException {
		Workbook workbook = parse(input, type, null);
		ComplexType resolve = (ComplexType) EAIResourceRepository.getInstance().resolve(typeId);
		ComplexContent newInstance = resolve.newInstance();
		for (Element<?> child : TypeUtils.getAllChildren(resolve)) {
			Value<String> property = child.getProperty(AliasProperty.getInstance());
			String sheetName = property == null ? child.getName() : property.getValue();
			if (sheetName.equals("*")) {
				sheetName = ".*";
			}
			newInstance.set(child.getName(), toObject(typeId, workbook, sheetName, true, null, null, null, rotate, includeEmptyResults, useHeaders, validateHeaders, true));
		}
		return newInstance;
	}
	
	@SuppressWarnings("unchecked")
	@WebResult(name = "marshalled")
	public InputStream marshal(@WebParam(name = "data") Object data, @WebParam(name = "excelType") FileType type, @WebParam(name = "useHeaders") Boolean useHeaders) throws IOException, ParseException {
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
		binding.marshal(output, (ComplexContent) data);
		return new ByteArrayInputStream(output.toByteArray());
	}
	
	@SuppressWarnings("resource")
	@WebResult(name = "results")
	public List<Object> toObject(@NotNull @WebParam(name = "typeId") String typeId, @WebParam(name = "workbook") Workbook workbook, @NotNull @WebParam(name = "sheet") String sheetName, @WebParam(name = "useRegexForSheet") Boolean useRegex, @WebParam(name = "fromRow") Integer fromRow, @WebParam(name = "toRow") Integer toRow, @WebParam(name = "columnsToIgnore") List<Integer> columnsToIgnore, @WebParam(name = "rotate") Boolean rotate, @WebParam(name = "includeEmptyResults") Boolean includeEmptyResults, @WebParam(name = "useHeaders") Boolean useHeaders, @WebParam(name = "validateHeaders") Boolean validateHeaders, @WebParam(name = "trim") Boolean trim) throws IOException, ParseException {
		// for backwards compatibility we trim by default
		if (trim == null) {
			trim = true;
		}
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
			if (useHeaders != null && useHeaders && ((fromRow == null && row == 0) || (fromRow != null && row == fromRow))) {
				if (validateHeaders != null && validateHeaders) {
					int elementCounter = 0;
					for (int column = 0; column < matrix.get(row).size(); column++) {
						if (columnsToIgnore == null || !columnsToIgnore.contains(column)) {
							// we ignore elements beyond the ones we can map, this could be empty cells or uninteresting data
							if (elementCounter >= children.size()) {
								break;
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
			for (int column = 0; column < matrix.get(row).size(); column++) {
				if (columnsToIgnore == null || !columnsToIgnore.contains(column)) {
					// we ignore elements beyond the ones we can map, this could be empty cells or uninteresting data
					if (elementCounter >= children.size()) {
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
