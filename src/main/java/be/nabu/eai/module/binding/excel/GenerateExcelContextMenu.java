package be.nabu.eai.module.binding.excel;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Sheet;

import be.nabu.eai.developer.MainController;
import be.nabu.eai.developer.api.EntryContextMenuProvider;
import be.nabu.eai.developer.managers.util.SimpleProperty;
import be.nabu.eai.developer.managers.util.SimplePropertyUpdater;
import be.nabu.eai.developer.util.EAIDeveloperUtils;
import be.nabu.eai.module.types.structure.StructureManager;
import be.nabu.eai.repository.api.Entry;
import be.nabu.eai.repository.resources.RepositoryEntry;
import be.nabu.libs.property.api.Property;
import be.nabu.libs.types.SimpleTypeWrapperFactory;
import be.nabu.libs.types.api.SimpleType;
import be.nabu.libs.types.base.SimpleElementImpl;
import be.nabu.libs.types.structure.DefinedStructure;
import be.nabu.utils.excel.ExcelParser;
import be.nabu.utils.excel.FileType;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.scene.control.MenuItem;
import javafx.scene.control.Menu;

public class GenerateExcelContextMenu implements EntryContextMenuProvider {

	@Override
	public MenuItem getContext(Entry entry) {
		if (!entry.isLeaf() && !entry.isNode()) {
			Menu menu = new Menu("Generate Model");
			
			MenuItem item = new MenuItem("From Excel");
			item.addEventHandler(ActionEvent.ANY, new EventHandler<ActionEvent>() {
				@Override
				public void handle(ActionEvent arg0) {
					Set<Property<?>> properties = new LinkedHashSet<Property<?>>();
					SimpleProperty<File> content = new SimpleProperty<File>("Excel", File.class, true);
					content.setInput(true);
					properties.add(content);
					properties.add(new SimpleProperty<FileType>("Type", FileType.class, false));
					final SimplePropertyUpdater updater = new SimplePropertyUpdater(true, properties);
					EAIDeveloperUtils.buildPopup(MainController.getInstance(), updater, "Generate structure from excel", new EventHandler<ActionEvent>() {
						@SuppressWarnings("resource")
						@Override
						public void handle(ActionEvent arg0) {
							FileType type = updater.getValue("Type");
							File excel = updater.getValue("Excel");
							if (type == null) {
								if (excel.getName().endsWith(".xls")) {
									type = FileType.XLS;
								}
								else {
									type = FileType.XLSX;
								}
							}
							try {
								InputStream stream = new BufferedInputStream(new FileInputStream(excel));
								try {
									ExcelParser excelParser = new ExcelParser(stream, type, null);
									for (String sheetName : excelParser.getSheetNames(false)) {
										String name = stringifyName(sheetName);
										// if we already have a child with this name, we are assuming you are regenerating, rather than generating
										if (entry.getChild(name) != null) {
											continue;
										}
										Sheet sheet = excelParser.getSheet(sheetName, false);
										List<List<Object>> matrix = excelParser.matrix(sheet);
										// we assume the first row is a header row, it should contain the names of the fields
										// we assume the second row is a content row, it should contain variables with correct types
										if (matrix.size() >= 2) {
											DefinedStructure structure = new DefinedStructure();
											structure.setName(name);
											for (int i = 0; i < matrix.get(0).size(); i++) {
												String label = matrix.get(0).get(i) == null ? null : matrix.get(0).get(i).toString();
												if (label != null) {
													Object value = null;
													if (i <= matrix.get(1).size() - 1) {
														value = matrix.get(1).get(i);
													}
													SimpleType resultingType = value == null 
														? SimpleTypeWrapperFactory.getInstance().getWrapper().wrap(String.class)
														: SimpleTypeWrapperFactory.getInstance().getWrapper().wrap(value.getClass());
													structure.add(new SimpleElementImpl(stringifyName(label), resultingType, structure));
												}
											}
											structure.setId(entry.getId() + "." + name);
											
											// only continue if we have a next element
											if (structure.iterator().hasNext()) {
												StructureManager manager = new StructureManager();
												RepositoryEntry repositoryEntry = ((RepositoryEntry) entry).createNode(name, manager, true);
												manager.saveContent(repositoryEntry, structure);
												MainController.getInstance().getRepositoryBrowser().refresh();
											}
										}
									}
								}
								finally {
									stream.close();
								}
							}
							catch (Exception e) {
								MainController.getInstance().notify(e);
							}
						}

						private String stringifyName(String sheetName) {
							StringBuilder name = new StringBuilder();
							for (String part : sheetName.split("[^\\w]+")) {
								if (name.toString().isEmpty()) {
									name.append(part.toLowerCase());
								}
								else {
									name.append(part.substring(0, 1).toUpperCase()).append(part.substring(1).toLowerCase());
								}
							}
							return name.toString();
						}
					});
				}
			});
			
			menu.getItems().add(item);
			return menu;
		}
		return null;
	}

}
