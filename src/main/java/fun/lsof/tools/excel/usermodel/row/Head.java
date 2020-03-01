package fun.lsof.tools.excel.usermodel.row;


/**
 * excel 头部对象类.
 *
 * @author jerry
 * @date 2017 -06-16 18:50:56
 */
public class Head
{
	/**
	 * 在 excel 列中的下标.
	 */
	int index;
	/**
	 * 下标对应的单元格文字.
	 */
	String text;

	/**
	 * Instantiates a new Head.
	 *
	 * @param index the index
	 * @param text  the text
	 * @author jerry
	 * @date 2017 -06-16 18:49:51
	 */
	public Head(int index, String text){
		super();
		this.index = index;
		this.text = text;
	}

	public int getIndex(){return index;}
	public void setIndex(int index){this.index = index;}
	public String getText(){return text;}
	public void setText(String text){this.text = text;}
	
	@Override
	public String toString(){
		return "{index"+":"+index+","+"text:"+text+"}";
	}
}
