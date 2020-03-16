import java.io.IOException;
import java.net.ServerSocket;

/**
 * @author chenningshui
 * @date 2020/3/16 15:27
 */
public class TestThread extends Thread{
    //volatile修饰符用来保证其它线程读取的总是该变量的最新的值
    public volatile boolean exit=false;

    @Override
    public void run(){
        try {
            ServerSocket serverSocket = new ServerSocket(8080);
            while (!exit){
                System.out.println("阻塞等待客户端消息中。。。。。");
                serverSocket.accept();//阻塞等待客户端消息
            }
        }catch (IOException e){
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        TestThread t =new TestThread();
        t.start();
        System.out.println("线程已经启动了");
//        try{
//            Thread.sleep(5000);
//        }catch (Exception e){}
        t.exit=true;//修改标志位，退出线程
    }
}
