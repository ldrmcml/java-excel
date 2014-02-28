package test;

public class test {
	public String pageURL="/action";;
	public int currentPage=1;
	public static void main(String[] args){
		test t=new test();
		String nextClick=" onclick='ajaxpage('"+t.pageURL+ "?currentPage=" + (t.currentPage + 1)+"');return false;'";
		System.out.println(nextClick);
	}

}
