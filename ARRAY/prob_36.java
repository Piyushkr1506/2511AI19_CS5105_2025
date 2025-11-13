package Arrays;

import java.util.Arrays;
import java.util.Scanner;

/*Find the repeating and missing number

Problem Statements: Given an integer array nums of size n containing values from [1, n] and each value appears exactly once in the array, 
except for A, which appears twice and B which is missing.
Return the values A and B, as an array of size 2, where A appears in the 0-th index and B in the 1st index.
Note: You are not allowed to modify the original array.

Example 1:
Input: nums = [3, 5, 4, 1, 1]
Output: [1, 2]
Explanation: 1 appears two times in the array and 2 is missing from nums

 */
public class prob_36 {


    static void MathSolution(int arr[],int n){
        //s-sn
        //s2-s2n
        long sn = (n*(n+1))/2;
        long s2n = (n*(n+1)*(2*n+1))/6;
        long s=0;
        long s2=0;
        for(int i=0;i<n;i++){
            s+=arr[i];
            s2+=arr[i]*arr[i];
        }
        long var1= s-sn;
        long var2 =s2-s2n;
        var2=var2/var1;
        long first = (var1+var2)/2;
        long second = first-var1;
        System.out.println(first+" "+second);
        
    }
    static int[] findMissAndRepNumber(int arr[],int n){
        int hashArr[] = new int[n+1];
        int repeating=-1;
        int missing=-1;
        for(int i=0;i<n;i++){
            hashArr[arr[i]]++;
        }
        for(int i=1;i<n+1;i++){
            if(hashArr[i]>=2){
                repeating=i;
            }
            if(hashArr[i]==0){
                missing=i;
            }
        }
        return new int[]{repeating,missing};

    }
    public static void main(String[] args) {
         Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array:");
        int N = sc.nextInt();
        int arr[] = new int[N];
        for(int i=0;i<N;i++){
            arr[i]= sc.nextInt();
        }
        int ans[] = findMissAndRepNumber(arr,N);
        System.out.println(Arrays.toString(ans));
        MathSolution(arr, N);
        sc.close();
    }
}
