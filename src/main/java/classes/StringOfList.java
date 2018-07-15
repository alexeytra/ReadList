package classes;

public class StringOfList {
    String cipher;
    String num;
    String content;

    public String getCipher() {
        return cipher;
    }

    public void setCipher(String cipher) {
        this.cipher = cipher;
    }

    public String getNum() {
        return num;
    }

    public void setNum(String num) {
        this.num = num;
    }

    public String getContent() {
        return content;
    }

    public void setContent(String content) {
        this.content = content;
    }

    public StringOfList(String cipher, String num, String content) {
        this.cipher = cipher;
        this.num = num;
        this.content = content;
    }
}
