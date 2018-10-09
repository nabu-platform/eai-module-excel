package be.nabu.eai.module.binding.excel;

import java.nio.charset.Charset;
import java.util.List;

import be.nabu.eai.module.rest.api.BindingProvider;
import be.nabu.libs.types.api.ComplexType;
import be.nabu.libs.types.binding.api.MarshallableBinding;
import be.nabu.libs.types.binding.api.UnmarshallableBinding;
import be.nabu.libs.types.binding.excel.ExcelBinding;
import be.nabu.utils.mime.api.Header;
import be.nabu.utils.mime.impl.MimeUtils;

public class ExcelBindingProvider implements BindingProvider {

	@Override
	public UnmarshallableBinding getUnmarshallableBinding(ComplexType type, Charset charset, Header... headers) {
		return null;
	}

	@Override
	public MarshallableBinding getMarshallableBinding(ComplexType type, Charset charset, Header... headers) {
		List<String> acceptedContentTypes = MimeUtils.getAcceptedContentTypes(headers);
		if (acceptedContentTypes != null && acceptedContentTypes.contains("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
			return getBinding(type, charset, headers);
		}
		return null;
	}
	
	private ExcelBinding getBinding(ComplexType type, Charset charset, Header... headers) {
		ExcelBinding excelBinding = new ExcelBinding();
		return excelBinding;
	}

	@Override
	public String getContentType(MarshallableBinding binding) {
		return "text/csv";
	}

}
