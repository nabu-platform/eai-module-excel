package be.nabu.eai.module.binding.excel;

import org.apache.poi.ss.usermodel.Workbook;

import be.nabu.libs.types.api.MarshalRuleProvider;

public class MarshalRuleProviderImpl implements MarshalRuleProvider {

	@Override
	public MarshalRule getMarshalRule(Class<?> clazz) {
		return Workbook.class.isAssignableFrom(clazz) ? MarshalRule.NEVER : null;
	}

}
