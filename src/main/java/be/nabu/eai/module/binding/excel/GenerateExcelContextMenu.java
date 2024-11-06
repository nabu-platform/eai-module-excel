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

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.UUID;

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
import be.nabu.libs.resources.api.ReadableResource;
import be.nabu.libs.resources.file.FileItem;
import be.nabu.libs.types.SimpleTypeWrapperFactory;
import be.nabu.libs.types.api.ComplexContent;
import be.nabu.libs.types.api.Element;
import be.nabu.libs.types.api.SimpleType;
import be.nabu.libs.types.base.SimpleElementImpl;
import be.nabu.libs.types.base.ValueImpl;
import be.nabu.libs.types.properties.LabelProperty;
import be.nabu.libs.types.properties.MinOccursProperty;
import be.nabu.libs.types.properties.PrimaryKeyProperty;
import be.nabu.libs.types.structure.DefinedStructure;
import be.nabu.libs.types.structure.Structure;
import be.nabu.libs.types.structure.StructureInstance;
import be.nabu.utils.excel.ExcelParser;
import be.nabu.utils.excel.FileType;
import be.nabu.utils.io.IOUtils;
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
							List<Structure> structures = parseExcel(type, new FileItem(null, excel, false), false);
							for (Structure structure : structures) {
								// if we already have a child with this name, we are assuming you are regenerating, rather than generating
								if (entry.getChild(structure.getName()) != null) {
									continue;
								}
								try {
									((DefinedStructure) structure).setId(entry.getId() + "." + structure.getName());
									StructureManager manager = new StructureManager();
									RepositoryEntry repositoryEntry = ((RepositoryEntry) entry).createNode(structure.getName(), manager, true);
									manager.saveContent(repositoryEntry, structure);
									MainController.getInstance().getRepositoryBrowser().refresh();
								}
								catch (Exception e) {
									MainController.getInstance().notify(e);
								}
							}
						}

						
					});
				}
			});
			
			menu.getItems().add(item);
			return menu;
		}
		return null;
	}
	
	private static String stringifyName(String sheetName) {
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
	
	public static List<Structure> parseExcel(FileType type, ReadableResource excel, boolean forDatabase) {
		List<Structure> structures = new ArrayList<Structure>();
		try {
			InputStream stream = new BufferedInputStream(IOUtils.toInputStream(excel.getReadable()));
			try {
				ExcelParser excelParser = new ExcelParser(stream, type, null);
				for (String sheetName : excelParser.getSheetNames(false)) {
					String name = stringifyName(sheetName);
					Sheet sheet = excelParser.getSheet(sheetName, false);
					List<List<Object>> matrix = excelParser.matrix(sheet);
					// we assume the first row is a header row, it should contain the names of the fields
					// we assume the second row is a content row, it should contain variables with correct types
					if (matrix.size() >= 2) {
						DefinedStructure structure = new DefinedStructure();
						if (forDatabase) {
							structure.add(new SimpleElementImpl<UUID>("id", SimpleTypeWrapperFactory.getInstance().getWrapper().wrap(UUID.class), structure,
								new ValueImpl<Boolean>(PrimaryKeyProperty.getInstance(), true)));
						}
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
								// we always add the label attribute so you can rename the field but we can still link it to the original excel
								structure.add(new SimpleElementImpl(stringifyName(label), resultingType, structure,
									new ValueImpl<String>(LabelProperty.getInstance(), label)));
							}
						}
						// only continue if we have a next element
						if (structure.iterator().hasNext()) {
							structures.add(structure);
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
		return structures;
	}
	
	public static Map<Structure, List<ComplexContent>> parseExcelAsContent(FileType type, ReadableResource excel, boolean forDatabase) {
		Map<Structure, List<ComplexContent>> result = new HashMap<Structure, List<ComplexContent>>(); 
		try {
			InputStream stream = new BufferedInputStream(IOUtils.toInputStream(excel.getReadable()));
			try {
				ExcelParser excelParser = new ExcelParser(stream, type, null);
				for (String sheetName : excelParser.getSheetNames(false)) {
					String name = stringifyName(sheetName);
					Sheet sheet = excelParser.getSheet(sheetName, false);
					List<List<Object>> matrix = excelParser.matrix(sheet);
					// we assume the first row is a header row, it should contain the names of the fields
					// we assume the second row is a content row, it should contain variables with correct types
					if (matrix.size() >= 2) {
						DefinedStructure structure = new DefinedStructure();
						if (forDatabase) {
							structure.add(new SimpleElementImpl<UUID>("id", SimpleTypeWrapperFactory.getInstance().getWrapper().wrap(UUID.class), structure,
								new ValueImpl<Boolean>(PrimaryKeyProperty.getInstance(), true)));
						}
						structure.setName(name);
						List<String> names = new ArrayList<String>();
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
								// we always add the label attribute so you can rename the field but we can still link it to the original excel
								SimpleElementImpl element = new SimpleElementImpl(stringifyName(label), resultingType, structure,
									new ValueImpl<String>(LabelProperty.getInstance(), label));
								structure.add(element);
								names.add(element.getName());
							}
							else {
								names.add(null);
							}
						}
						// only continue if we have a next element
						if (structure.iterator().hasNext()) {
							List<ComplexContent> instances = new ArrayList<ComplexContent>();
							// loop over the rows, map it
							for (int i = 1; i < matrix.size(); i++) {
								StructureInstance instance = structure.newInstance();
								for (int j = 0; j < matrix.get(i).size(); j++) {
									String columnName = names.get(j);
									// we can only add the field if we have a name for it
									if (columnName != null) {
										Object object = matrix.get(i).get(j);
										Element<?> element = structure.get(columnName);
										// make sure the type is optional
										if (object == null) {
											element.setProperty(new ValueImpl<Integer>(MinOccursProperty.getInstance(), 0));
										}
										else {
											instance.set(columnName, object);
										}
									}
								}
								instances.add(instance);
							}
							result.put(structure, instances);
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
		return result;
	}

}
