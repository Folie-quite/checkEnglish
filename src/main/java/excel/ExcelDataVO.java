package excel;

/**
 * Author: Dreamer-1
 * Date: 2019-03-01
 * Time: 11:33
 * Description: 读取Excel时，封装读取的每一行的数据
 */
public class ExcelDataVO {

    /**
     * 中文
     */
    private String chinese;

    /**
     * 英文
     */
    private String english;

    /**
     * 回答
     */
    private String answer;

	/**
	 * @return the chinese
	 */
	public String getChinese() {
		return chinese;
	}

	/**
	 * @param chinese the chinese to set
	 */
	public void setChinese(String chinese) {
		this.chinese = chinese;
	}

	/**
	 * @return the english
	 */
	public String getEnglish() {
		return english;
	}

	/**
	 * @param english the english to set
	 */
	public void setEnglish(String english) {
		this.english = english;
	}

	/**
	 * @return the answer
	 */
	public String getAnswer() {
		return answer;
	}

	/**
	 * @param answer the answer to set
	 */
	public void setAnswer(String answer) {
		this.answer = answer;
	}
    
}
