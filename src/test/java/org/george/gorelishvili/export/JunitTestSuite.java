package org.george.gorelishvili.export;

import org.george.gorelishvili.export.common.TestCellBuilder;
import org.george.gorelishvili.export.common.TestSheetWriter;
import org.george.gorelishvili.export.writer.TestFastWriter;
import org.junit.runner.RunWith;
import org.junit.runners.Suite;

@RunWith(Suite.class)
@Suite.SuiteClasses({
 	TestCellBuilder.class,
	TestFastWriter.class,
	TestSheetWriter.class
})
public class JunitTestSuite {
}