#!/bin/sh

export JAVA_HOME=/opt/ibm/ibm-java-x86_64-60
export CLASS_PATH=$(echo lib/*.jar | tr ' ' ':')

$JAVA_HOME/bin/java -cp $(echo lib/*.jar | tr ' ' ':') com.ibm.custom.WordHelper samplet.dotx output.doc
