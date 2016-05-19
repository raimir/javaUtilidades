package xtr.util;

import javax.xml.bind.DatatypeConverter; //java 6+

public class Base64 {

	// encode
	public String encode2(String texto) {
		try {
			return DatatypeConverter.printBase64Binary(texto.getBytes("UTF-8"));
		} catch (Exception e) {
			e.printStackTrace();
			return texto;
		}
	}

	//decode
	public static String decode(String texto) throws Exception {
		try {
			byte[] decodeValor = DatatypeConverter.parseBase64Binary(texto);
			return new String(decodeValor, "UTF-8");
		} catch (Exception e) {
			e.printStackTrace();
			return texto;
		}
	}
}
