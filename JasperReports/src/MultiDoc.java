

public class MultiDoc
{

	private static Double docCount = null;		// Flag to indicate single or multi page
    private static String docRoot = null;		// First element in XML source
	private	static String docElement = null;	// Iterating reference
	private	static String docKey = null;		// Identifier for an iterating reference
	
	public void setDocCount(Double x)
	{
		docCount = x;
	}

	public void setDocRoot(String x)
	{
		docRoot = x;
	}

	public void setDocElement(String x)
	{
		docElement = x;
	}

	public void setDocKey(String x)
	{
		docKey = x;
	}

	/*---------------------------------------*/
	public Double getDocCount()
	{
		return docCount;
	}

	public String getDocRoot()
	{
		return docRoot;
	}

	public String getDocElement()
	{
		return docElement;
	}

	public String getDocKey()
	{
		return docKey;
	}

}