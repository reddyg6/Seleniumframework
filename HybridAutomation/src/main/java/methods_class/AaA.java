package methods_class;

import java.net.HttpURLConnection;
import java.net.URL;

public class AaA {
	public static void main(String[] args) {
		try {
			URL link = new URL("https://www.hsbc.co.uk/insurance/products/home/");
			HttpURLConnection httpConn = (HttpURLConnection) link.openConnection();
			httpConn.setConnectTimeout(2000);
			httpConn.connect();
			if (httpConn.getResponseCode() == 200) {
				System.out.println(link + " - " + httpConn.getResponseMessage());
			}
			if (httpConn.getResponseCode() == 404) {
				System.out.println(link + " - " + httpConn.getResponseMessage());
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
