package i18n.xlsx;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.nio.charset.Charset;
import java.nio.charset.CharsetEncoder;


public class UnquotingWriter extends OutputStreamWriter {
	
	private final char quoteChar = '"';
	private final char buff[] = new char[512];
	
	private boolean isQuoted = false;
	private boolean precededByQuote = false;

	public UnquotingWriter(OutputStream out, Charset cs) {
		super(out, cs);
		// TODO Auto-generated constructor stub
	}

	public UnquotingWriter(OutputStream out, CharsetEncoder enc) {
		super(out, enc);
		// TODO Auto-generated constructor stub
	}

	public UnquotingWriter(OutputStream out, String charsetName)
			throws UnsupportedEncodingException {
		super(out, charsetName);
		// TODO Auto-generated constructor stub
	}

	public UnquotingWriter(OutputStream out) {
		super(out);
		// TODO Auto-generated constructor stub
	}

	@Override
	public void write(int c) throws IOException {
		if (c == quoteChar) {
			isQuoted = !isQuoted;
			if (isQuoted) {
				if (precededByQuote) {
					super.write(c);
					precededByQuote = false;
				}
			} else {
				precededByQuote = true;
			}
			return;
		} else {
			precededByQuote = false;
		}
		super.write(c);
	}	

	@Override
	public void write(char[] cbuf, int off, int len) throws IOException {
		int buffLength = 0;
		for(int i = 0; i < len; i++) {
			final char c = cbuf[off + i];
			buffLength = writeUnquoted(buffLength, c);
		}
		super.write(buff, 0, buffLength);
	}

	@Override
	public void write(String str, int off, int len) throws IOException {
		int buffLength = 0;
		for(int i = 0; i < len; i++) {
			final char c = str.charAt(off + i);
			buffLength = writeUnquoted(buffLength, c);
		}
		super.write(buff, 0, buffLength);
	}

	private int writeUnquoted(final int initialOffset, final char c) throws IOException {
		int offset = initialOffset;
		if (c == quoteChar) {
			isQuoted = !isQuoted;
			if (isQuoted) {
				if (precededByQuote) {
					buff[offset++] = c;
					precededByQuote = false;
				}
			} else {
				precededByQuote = true;
			}
			return offset;
		} else {
			precededByQuote = false;
		}
		buff[offset++] = c;
		if(offset >= buff.length) {
			super.write(buff, 0, offset);
			offset = 0;
		}
		return offset;
	}
	
	public void reset() {
		this.isQuoted = false;
		this.precededByQuote = false;
	}
}
