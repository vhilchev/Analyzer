package org.analyzer;

import javax.swing.*;

/**
 * Created by vhilc on 5/25/2017.
 */
public class Main {
    public static void main(String[] args) {
        JFrame frame = new JFrame("Analyzer");
        frame.setContentPane(new Dashboard().panel1);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(450,150);
        frame.setResizable(true);
        frame.setVisible(true);
    }
}
