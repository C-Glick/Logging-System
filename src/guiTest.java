import java.awt.*;

import javax.swing.*;

public class guiTest extends JFrame {

    public guiTest() throws HeadlessException {
        final JPanel panel = new JPanel();
        panel.setBorder(BorderFactory.createLineBorder(Color.red));
        panel.setPreferredSize(new Dimension(800, 600));

        final JScrollPane scroll = new JScrollPane(panel);
        panel.setLayout(null);
        
        JLabel lblHi = new JLabel("hi");
        lblHi.setBounds(72, 131, 46, 14);
        panel.add(lblHi);
        
        JLabel lblHello = new JLabel("hello");
        lblHello.setBounds(463, 403, 46, 14);
        panel.add(lblHello);

        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        getContentPane().setLayout(new BorderLayout());
        getContentPane().add(scroll, BorderLayout.CENTER);
        setSize(609, 512);
        setVisible(true);
    }

    public static void main(final String[] args) throws Exception {
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                new guiTest().setVisible(true);
            }
        });
    }
}