package Arrays;

import java.util.HashMap;
import java.util.Map;
import java.util.Scanner;

/*Majority Element-1

Problem Statements: Given an integer array nums of size n, return the majority element of the array.

The majority element of an array is an element that appears more than n/2 times in the array. The array is guaranteed to have a majority element.

Example 1:
Input: nums = [7, 0, 0, 1, 7, 7, 2, 7, 7]

Output: 7

Explanation: The number 7 appears 5 times in the 9 sized array
 */
public class prob_10 {
    public static void main(String[] args) {

         Scanner sc = new Scanner(System.in);
        System.out.println("Enter the size of Array:");
        int N = sc.nextInt();
        int arr[] = new int[N];
        for(int i=0;i<N;i++){
            arr[i]= sc.nextInt();
        }
        HashMap<Integer,Integer> hm = new HashMap<>();
        for(int i=0;i<N;i++){
            if(!hm.containsKey(arr[i])){
                hm.put(arr[i],1);
            }else{
                hm.put(arr[i],hm.get(arr[i])+1);
            }
        }
        for (Map.Entry<Integer,Integer> mapElement : hm.entrySet()) {
            int key = mapElement.getKey();
            int value = mapElement.getValue();
            if(value > N/2){
                System.out.println("The majority element in the given array is :"+ key);
            }
        }
        sc.close();
    }
    
}
