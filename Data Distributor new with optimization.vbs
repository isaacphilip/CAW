O p t i o n   E x p l i c i t  
 D i m   o b j E x c e l ,   o b j P D H Q ,   s t r E x c e l P a t h ,   s t r P D H Q ,   o b j S h e e t ,   o b j S h e e t 2 ,   r o w s ,   c o l ,   r o w ,   w o r d s ,   a ,   x  
 D i m   s t r e s e t u p ,   o b j e s e t u p ,   o b j S h e e t e ,   o b j S h e l l ,   s t r C u r D i r ,   r o w p  
 D i m   f s o ,   o b j c o n f i g ,   f i l e n a m e ,   f ,   g ,   n a m e ,   o b j T e x t F i l e  
  
 S e t   o b j S h e l l   =   C r e a t e O b j e c t ( " W s c r i p t . S h e l l " )  
  
 S e t   f s o   =   C r e a t e O b j e c t ( " S c r i p t i n g . F i l e S y s t e m O b j e c t " )  
 F o r   E a c h   g   i n   f s o . G e t F o l d e r ( " C : \ B M S   A u t o m a t i o n \ l o g " ) . F i l e s  
     n a m e   =   L C a s e ( g . N a m e )  
     I f   f s o . G e t E x t e n s i o n N a m e ( n a m e )   =   " t x t "   T h e n  
         S e t   o b j T e x t F i l e   =   f s o . O p e n T e x t F i l e   ( " C : \ B M S   A u t o m a t i o n \ l o g \ " &   g . N a m e ,   8 ,   T r u e )  
     E n d   I f  
 N e x t  
 s t r C u r D i r   =   o b j S h e l l . C u r r e n t D i r e c t o r y    
 f s o . C o p y F i l e   " \ \ U S H P W B M S F S P 0 0 2 . O N E . A D S . B M S . C O M \ s h a r e d 0 2 \ C A A \ C A A G r o u p s \ c o n f i g \ * " ,   " C : \ T e m p \ " , T r u e  
 W S c r i p t . S l e e p   5 0 0 0  
 S e t   o b j c o n f i g   =   C r e a t e O b j e c t ( " E x c e l . A p p l i c a t i o n " )  
 o b j c o n f i g . W o r k B o o k s . O p e n   ( " C : \ T e m p \ c o n f i g . x l s x " )  
 S e t   o b j S h e e t 2   =   o b j c o n f i g . A c t i v e W o r k b o o k . W o r k s h e e t s ( 1 )  
  
 S e t   o b j E x c e l   =   C r e a t e O b j e c t ( " E x c e l . A p p l i c a t i o n " )  
 o b j E x c e l . W o r k B o o k s . O p e n   o b j c o n f i g . C e l l s ( 1 , 2 ) . V a l u e  
 S e t   o b j S h e e t   =   o b j E x c e l . A c t i v e W o r k b o o k . W o r k s h e e t s ( 1 )  
 s t r e s e t u p   =   o b j c o n f i g . C e l l s ( 2 , 2 ) . V a l u e  
 f i l e n a m e   =   o b j c o n f i g . C e l l s ( 3 , 2 ) . V a l u e  
  
  
  
 ' W s c r i p t . E c h o   s t r e s e t u p  
  
 D I M   I E  
 D I M   u r l s  
 S e t   I E   =   C r e a t e O b j e c t ( " I n t e r n e t E x p l o r e r . A p p l i c a t i o n " )  
  
 r o w s   =   1  
 w o r d s   =   0  
 c o l   =   1  
 r o w   =   1  
 r o w p =   1  
  
 S e t   o b j e s e t u p   =   C r e a t e O b j e c t ( " E x c e l . A p p l i c a t i o n " )  
 o b j e s e t u p . W o r k B o o k s . O p e n   s t r e s e t u p  
 S e t   o b j S h e e t e   =   o b j e s e t u p . A c t i v e W o r k b o o k . W o r k s h e e t s ( 1 )  
  
 D o   U n t i l   o b j E x c e l . C e l l s ( r o w s , 1 ) . V a l u e   =     " "  
   D o   U n t i l   o b j c o n f i g . C e l l s ( r o w , 1 ) . V a l u e   =     " "  
      
      
     I f   ( L e f t ( o b j E x c e l . C e l l s ( r o w s , 1 ) . V a l u e , 2 0 )   =     L e f t ( o b j c o n f i g . C e l l s ( r o w , 1 ) . V a l u e , 2 0 ) )   T h e n  
       x   =   S p l i t ( o b j E x c e l . C e l l s ( r o w s , 1 ) . V a l u e , " [ " )  
       a   =   S p l i t ( x ( 1 ) , " ] " )  
       I f   a ( 0 )   =   o b j c o n f i g . C e l l s ( r o w , 5 ) . V a l u e   T h e n  
         I f ( o b j c o n f i g . C e l l s ( r o w , 4 ) . V a l u e   =   1 )   T h e n  
           o b j S h e e t e . C e l l s ( r o w s ,   1 ) . V a l u e   =   o b j E x c e l . C e l l s ( r o w s , 1 ) . V a l u e  
           o b j S h e e t e . C e l l s ( r o w s ,   2 ) . V a l u e   =   o b j S h e e t . C e l l s ( r o w s ,   2 ) . V a l u e  
           o b j e s e t u p . A c t i v e W o r k b o o k . S a v e  
         E n d   I f  
         I E . V i s i b l e   =   1  
         S e t   o b j P D H Q   =   C r e a t e O b j e c t ( " E x c e l . A p p l i c a t i o n " )  
         o b j P D H Q . W o r k B o o k s . O p e n   o b j c o n f i g . C e l l s ( r o w , 2 ) . V a l u e  
         S e t   o b j S h e e t 2   =   o b j P D H Q . A c t i v e W o r k b o o k . W o r k s h e e t s ( 1 )  
         I E . N a v i g a t e   o b j S h e e t . C e l l s ( r o w s ,   2 ) . V a l u e  
         W h i l e   I E . R e a d y S t a t e   < >   4  
             W S c r i p t . S l e e p   1 0 0 0  
         W e n d  
         s e t   u r l s   =   i e . d o c u m e n t . a l l . t a g s ( " a " )  
         c o l   =   1  
         W S c r i p t . S l e e p   1 0 0 0  
         D o   U n t i l   o b j S h e e t 2 . C e l l s ( r o w p ,   1 ) . V a l u e   =   " "  
           r o w p   =   r o w p   +   1  
         L o o p  
         D o   U n t i l   w o r d s   >   1 5 0   O R   I E . d o c u m e n t . g e t E l e m e n t s B y T a g N a m e ( " f o n t " ) . I t e m ( w o r d s ) . I n n e r T e x t   =     " - "  
           I f   c o l   =   4   T h e n  
             o b j S h e e t 2 . C e l l s ( r o w p ,   c o l ) . V a l u e   = M i d ( u r l s ( 2 ) . i n n e r H T M L , 7 , 9 )  
             o b j P D H Q . A c t i v e W o r k b o o k . S a v e  
           ' W s c r i p t . E c h o   o b j S h e e t 2 . C e l l s ( r o w p ,   c o l ) . V a l u e  
             c o l   =   c o l + 1  
           E l s e  
             I f ( w o r d s   <   4   O R   w o r d s   >   o b j c o n f i g . C e l l s ( r o w , 3 ) . V a l u e )   T h e n  
 	       o b j S h e e t 2 . C e l l s ( r o w p ,   c o l ) . V a l u e   =   I E . d o c u m e n t . g e t E l e m e n t s B y T a g N a m e ( " f o n t " ) . I t e m ( w o r d s ) . I n n e r T e x t  
 	       S e t   f   =   f s o . O p e n T e x t F i l e ( f i l e n a m e )  
 	       D o   U n t i l   f . A t E n d O f S t r e a m  
 	         ' W s c r i p t . E c h o   f . R e a d L i n e  
                 o b j S h e e t 2 . C e l l s ( r o w p ,   c o l ) . V a l u e   =   T r i m ( R e p l a c e ( o b j S h e e t 2 . C e l l s ( r o w p ,   c o l ) . V a l u e , f . R e a d L i n e , "   " ) )  
 	         ' o b j P D H Q . A c t i v e W o r k b o o k . S a v e  
               L o o p  
 	       f . C l o s e  
               ' o b j S h e e t 2 . C e l l s ( r o w p ,   c o l ) . V a l u e   =   T r i m ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( R e p l a c e ( I E . d o c u m e n t . g e t E l e m e n t s B y T a g N a m e ( " f o n t " ) . I t e m ( w o r d s ) . I n n e r T e x t , " < " , "   " ) , " > " , "   " ) , " ( U I D = " , "   " ) , " ) " ,   "   " ) , " o p t i o n : " , "   " ) , " A c t i o n : " , "   " ) , " S e l e c t   a p p l i c a t i o n : " , "   " ) , " S e l e c t   t h e   a p p r o p r i a t e   r o l e : " , "   " ) , " =   ( " ,   "   " ) , " R o l e s : " , "   " ) , " E n v i r o n m e n t : " , "   " ) , " S i t e : " , "   " ) , " = " , "   " ) , " D i g i t a l   C e r t i f i c a t e   T y p e : " , "   " ) , " S e l e c t   a   P r o j e c t :   " , "   " ) , " S e l e c t   G r o u p / R o l e " , "   " ) , " S e l e c t   r o l e / o p t i o n : " , "   " ) , " M S   P r o j e c t   P r o   N e e d e d ? : " , "   " ) , " R o l e : " , "   " ) , " S p e c i f y   R o l e   ( E D C   T e s t   O n l y   : " , "   " ) , " S p e c i f y   R o l e   ( E D C   : " , "   " ) , " A l l i a n c e   G r o u p s :   ( h o l d   C T R L   t o   m u l t i - s e l e c t " , "   " ) , " S p e c i f y " , "   " ) , " c o u n t r y : " , "   " ) , " D i v i s i o n : " , "   " ) , " C o u n t r y : " , "   " ) , " S e l e c t " , "   " ) , " A p p l i c a t i o n ( s   : " , "   " ) , " T y p e : " , "   " ) )  
               ' o b j P D H Q . A c t i v e W o r k b o o k . S a v e  
 	       o b j T e x t F i l e . W r i t e L i n e ( N o w ( ) )  
               o b j T e x t F i l e . W r i t e L i n e ( " W r o t e   t o   d a t a s h e e t :   " &   o b j S h e e t 2 . C e l l s ( r o w p ,   c o l ) . V a l u e )  
               c o l   =   c o l + 1  
             E n d   I f  
             w o r d s =   w o r d s   + 1  
           E n d   I f  
         L o o p  
         r o w p = r o w p + 1  
         o b j P D H Q . A c t i v e W o r k b o o k . S a v e  
         o b j P D H Q . A c t i v e W o r k b o o k . C l o s e  
         o b j P D H Q . A p p l i c a t i o n . Q u i t  
         o b j P D H Q . Q u i t        
         ' r o w s   =   r o w s   +   1  
       E n d   I f  
     E n d   I f  
     r o w   =   r o w   + 1  
     w o r d s   =   0  
     r o w p   =   1  
   L o o p  
   r o w s   =   r o w s   +   1  
   r o w   =   1  
 L o o p  
 o b j e s e t u p . A c t i v e W o r k b o o k . S a v e  
 o b j e s e t u p . A c t i v e W o r k b o o k . C l o s e  
 o b j e s e t u p . A p p l i c a t i o n . Q u i t  
 o b j e s e t u p . Q u i t  
 o b j E x c e l . A c t i v e W o r k b o o k . C l o s e  
 o b j E x c e l . A p p l i c a t i o n . Q u i t  
 o b j E x c e l . Q u i t  
 o b j c o n f i g . A c t i v e W o r k b o o k . C l o s e  
 o b j c o n f i g . A p p l i c a t i o n . Q u i t  
 o b j c o n f i g . Q u i t  
 o b j S h e l l . R u n ( " t a s k k i l l   / i m   E X C E L . e x e " ) ,   1 ,   T R U E  
 f s o . D e l e t e F i l e   " C : \ T e m p \ c o n f i g . x l s x " ,   T r u e    
 M s g b o x   " c o m p l e t e d " 