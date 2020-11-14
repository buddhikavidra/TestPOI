import java.io.IOException;

public class Reader {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		ReadFile r = new ReadFile();
		String val = r.getMapData("prm1");
		System.out.println("Value of the keyword () is- "+val);

	}

}
