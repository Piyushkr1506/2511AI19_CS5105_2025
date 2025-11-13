package Arrays;

import java.util.Scanner;

/*Print the matrix in spiral manner

Problem Statements: Given an M * N matrix, print the elements in a clockwise spiral manner. Return an array with the elements in the order of their appearance when printed in a spiral manner.

Example 1:
Input: matrix = [[1, 2, 3], [4 ,5 ,6], [7, 8, 9]]

Output: [1, 2, 3, 6, 9, 8, 7, 4, 5]

Explanation: The elements in the spiral order are 1, 2, 3 -> 6, 9 -> 8, 7 -> 4, 5
 
*/
public class prob_28 {

    static void PrintSpiral(int[][] mat,int n,int m)
    {
        int top=0;
        int left=0;
        int right=m-1;
        int bottom=n-1;
        while(left<=right && top<=bottom){
            for(int i=left;i<=right;i++){
            System.out.print(mat[top][i]+" ");
        }
        top++;
        for(int j=top;j<=bottom;j++){
            System.out.print(mat[j][right]+" ");
        }
        right--;
        if(top<=bottom){
          for(int k=right;k>=left;k--){
            System.out.print(mat[bottom][k]+" ");
        }
        bottom--;  
        }
        if(left<=right){
           for(int l=bottom;l>=top;l--){
            System.out.print(mat[l][left]+" ");
                }
        left++; 
            }
        }    
    }
    public static void main(String[] args) {
         Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array N M:");
        int N = sc.nextInt();
        int M =sc.nextInt();
        int matrix[][] = new int[N][M];
        System.out.println("enter the elements of matrix:");
        for(int i=0;i<N;i++){
            for(int j=0;j<M;j++){
                matrix[i][j]= sc.nextInt();
            }
        }

        PrintSpiral(matrix,N,M);
        sc.close();
    }
}
