package filemerger;


import java.awt.BorderLayout;
import java.awt.Point;
import javax.swing.GroupLayout;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JTextPane;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author mmercado
 */
public class View {

    private JFrame frame;
    private JButton btnSelect;
    private JButton btnExtract;
    private JTextPane selectPane;
    private JScrollPane scrollPane;
    private JProgressBar progressBar;

    public View(String title) {
        frame = new JFrame(title);
        frame.getContentPane().setLayout(new BorderLayout());
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(400, 284);
        frame.setLocation(new Point(500, 300));
        frame.setResizable(false);
        frame.setVisible(true);

        btnSelect = new JButton("Select Files");
        btnExtract = new JButton("Extract Files");
        selectPane = new JTextPane();
        scrollPane = new JScrollPane();
        progressBar = new JProgressBar();

        GroupLayout layout = new GroupLayout(frame.getContentPane());
        layout.setHorizontalGroup(
                layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                        .addGroup(layout.createSequentialGroup()
                                .addContainerGap()
                                .addGroup(layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                                        .addComponent(scrollPane, GroupLayout.Alignment.LEADING)
                                        .addComponent(progressBar)
                                )
                                .addContainerGap()
                        )
                        .addGroup(layout.createSequentialGroup()
                                .addGap(45, 45, 45)
                                .addComponent(btnSelect, GroupLayout.PREFERRED_SIZE, 132, GroupLayout.PREFERRED_SIZE)
                                .addGap(46, 46, 46)
                                .addComponent(btnExtract, GroupLayout.PREFERRED_SIZE, 132, GroupLayout.PREFERRED_SIZE)
                                .addContainerGap(45, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
                layout.createParallelGroup(GroupLayout.Alignment.LEADING)
                        .addGroup(GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                                .addContainerGap(GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addComponent(scrollPane, GroupLayout.PREFERRED_SIZE, 100, GroupLayout.PREFERRED_SIZE)
                                .addGap(41, 41, 41)
                                .addGroup(layout.createParallelGroup(GroupLayout.Alignment.BASELINE)
                                        .addComponent(btnSelect, GroupLayout.PREFERRED_SIZE, 50, GroupLayout.PREFERRED_SIZE)
                                        .addComponent(btnExtract, GroupLayout.PREFERRED_SIZE, 50, GroupLayout.PREFERRED_SIZE))
                                .addGap(20, 20, 20)
                                .addComponent(progressBar, GroupLayout.PREFERRED_SIZE, 24, GroupLayout.PREFERRED_SIZE)
                                .addContainerGap())
        );
        frame.getContentPane().setLayout(layout);

    }

    public JFrame getFrame() {
        return frame;
    }

    public void setFrame(JFrame frame) {
        this.frame = frame;
    }

    public JButton getBtnSelect() {
        return btnSelect;
    }

    public void setBtnSelect(JButton btnSelect) {
        this.btnSelect = btnSelect;
    }

    public JButton getBtnExtract() {
        return btnExtract;
    }

    public void setBtnExtract(JButton btnExtract) {
        this.btnExtract = btnExtract;
    }

    public JTextPane getSelectPane() {
        return selectPane;
    }

    public void setSelectPane(JTextPane selectPane) {
        this.selectPane = selectPane;
    }

    public JScrollPane getScrollPane() {
        return scrollPane;
    }

    public void setScrollPane(JScrollPane scrollPane) {
        this.scrollPane = scrollPane;
    }

    public JProgressBar getProgressBar() {
        return progressBar;
    }

    public void setProgressBar(JProgressBar progressBar) {
        this.progressBar = progressBar;
    }

}
