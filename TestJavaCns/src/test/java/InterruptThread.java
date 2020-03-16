/**
 * @author chenningshui
 * @date 2020/3/16 15:50
 */
public class InterruptThread extends Thread {

    @Override
    public void run(){
        super.run();
        for(int i=0;i<20000;i++)
        {
            //判断是否被中断
            if(Thread.currentThread().isInterrupted()){
                System.out.println("线程被中断");
                break;
            }
            System.out.println(i);
        }
    }

    public static void main(String[] args) {
        try {
            InterruptThread t = new InterruptThread();
            t.start();
            Thread.sleep(200);
            t.interrupt();
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
