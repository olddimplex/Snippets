package i18n.xlsx;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.Writer;

import common.IStringTransformer;


public class UnquotingStringTransformer implements IStringTransformer {

	@Override
	public String transform(String str) {
		if(str != null) {
			try(final ByteArrayOutputStream baos = new ByteArrayOutputStream();
				final Writer writer = new UnquotingWriter(baos)) {
				writer.write(str, 0, str.length());
				writer.flush();
				return baos.toString();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return str;
	}

}
