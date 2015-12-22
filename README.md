# Non-maven project

The apache poi modules do not have maven descriptor files which means the maven classloader can not automatically find them.

## Dependencies

In theory we should only define dependencies to apache poi itself, it will get the other dependencies.
However because the apache poi library does not contain a pom file describing these dependencies the maven classloader won't find them.

Instead we have opted to explicitly declare the dependencies in the pom file of this project, forcing them to be found.

## Packaging

The maven repository has a fallback in case a pom file is not found: it will use the name of the file to determine the properties: `groupId-artifactId-version.jar`.

However the artifacts as we get them from the central maven repository only have `artifactId-version.jar`. 
We still need to add the groupId or the repository will default to the group "com.example" which will make it unfindable for our dependencies.

To rename the dependencies, we update the assembly xml of our packager:

```xml
<fileSets>
	<fileSet>
		<directory>${project.build.directory}/lib</directory>
		<outputDirectory>/</outputDirectory>
		<includes>
			<include>*.*</include>
		</includes>
		<excludes>
			<exclude>poi-3.13.jar</exclude>
			<exclude>poi-ooxml-3.13.jar</exclude>
			<exclude>poi-ooxml-schemas-3.13.jar</exclude>
			<exclude>xmlbeans-2.6.0.jar</exclude>
		</excludes>
	</fileSet>
</fileSets>
<files>
	<file>
		<source>${project.build.directory}/lib/poi-3.13.jar</source>
		<destName>org.apache.poi-poi-3.13.jar</destName>
	</file>
	<file>
		<source>${project.build.directory}/lib/poi-ooxml-3.13.jar</source>
		<destName>org.apache.poi-poi-ooxml-3.13.jar</destName>
	</file>
	<file>
		<source>${project.build.directory}/lib/poi-ooxml-schemas-3.13.jar</source>
		<destName>org.apache.poi-poi-ooxml-schemas-3.13.jar</destName>
	</file>
	<file>
		<source>${project.build.directory}/lib/xmlbeans-2.6.0.jar</source>
		<destName>org.apache.xmlbeans-xmlbeans-2.6.0.jar</destName>
	</file>
</files>
```

We have excluded the files that we want to rename from the original fileset.
Then we have added those files back again with a fixed name.

An alternative would've been to explicitly state dependencies to the known "default" groupId but this would make this project very hard to build as default maven would not find them.