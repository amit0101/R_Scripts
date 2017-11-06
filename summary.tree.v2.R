
# Function to create excel summary of rpart tree outputs
# Requires library XLConnect

summary.tree<-function(x,data,numerator,denominator,id) {
  
  sink("temp.txt")
  print(x)
  sink()
  
  y=readLines("temp.txt")
  y=y[grep("1)",y,fixed=TRUE)[1]:length(y)]
  y=gsub("< ","<",y)
  y=gsub("  "," ",y)
  
  z=data.frame(t(sapply(t(data.frame(y)),FUN=function(w) setdiff(unlist(strsplit(trimws(w)," ")),c("","*")))))
  z[,1]=gsub(")","",as.character(z[,1]))
  z[,2]=as.character(z[,2])
  z[,3:5]=sapply(z[,3:5],FUN=function(w) as.numeric(as.character(w)))
  colnames(z)=c("node","categories","count","deviance","y")
  rownames(z)=1:nrow(z)
  z$splitvar=apply(t(z[,2]),2,
                   FUN=function(w) unlist(strsplit(unlist(strsplit(unlist(strsplit(unlist(strsplit(unlist(strsplit(w,">=",fixed=TRUE)),"<=",fixed=TRUE)),">",fixed=TRUE)),"<",fixed=TRUE)),"=",fixed=TRUE))[1])
  z$catlist=apply(t(z[,2]),2,
                  FUN=function(w) paste("c('",gsub(",","','",unlist(strsplit(unlist(strsplit(unlist(strsplit(unlist(strsplit(unlist(strsplit(w,">=",fixed=TRUE)),"<=",fixed=TRUE)),">",fixed=TRUE)),"<",fixed=TRUE)),"=",fixed=TRUE))[2]),"')",sep=""))
  z$leaf=0
  z$leaf[grep("*",z,fixed=TRUE)]=1
  z$operator="%in%"
  z$operator[grep("<",z$categories)]="<"
  z$operator[grep(">",z$categories)]=">"
  z$operator[grep(">=",z$categories)]=">="
  z$operator[grep("<=",z$categories)]="<="
  
  parent <- function(x) {
    if (x[1] != 1)
      c(Recall(if (x %% 2 == 0L) x / 2 else (x - 1) / 2), x) else x
  }
  
  depth=ceiling(log2(max(as.numeric(z$node))))-1
  
  splits=data.frame(1)
  colnames(splits)=0
  subfamily=data.frame()
  
  for(level in 1:depth) {
    cat(paste("\nLevel",level,"tree structure\n"))
    parents=data.frame(unique(splits[,level])[!is.na(unique(splits[,level]))])
    subfamily=data.frame()
    for(i in 1:nrow(parents)) {
      if(length(grep(2*parents[i,],z$node))>0) {subfamily=rbind(subfamily,c(parents[i,],2*parents[i,]))}
      if(length(grep((2*parents[i,])+1,z$node))>0) {subfamily=rbind(subfamily,c(parents[i,],(2*parents[i,])+1))}
    }
    colnames(subfamily)[2]=as.character(level)
    subfamily=subfamily[order(subfamily[,2]),]
    splits=merge(splits,subfamily,by.x=colnames(splits)[level],by.y=colnames(subfamily)[1],all.x=TRUE)
    splits=splits[,as.character(0:level)]
    print(splits[,order(colnames(splits))])
  }
  
  for(i in ncol(splits):2) {
    splits=splits[order(splits[,i]),]
  }
  
  library(XLConnect)
  tstamp=gsub(" ","",gsub("[[:punct:]]","",as.character(Sys.time())))
  wb <- loadWorkbook(paste("treeSummary_",tstamp,".xlsx",sep=""), create = TRUE)
  createSheet(wb, name = "splits")
  cat("\nBuilding tree summary\n---------------------\n")
  
  for(i in 1:ncol(splits)) {
    for(node in unique(splits[,i])[!is.na(unique(splits[,i]))]) {
      start=min(grep(node,splits[,i]))
      end=max(grep(node,splits[,i]))
      if(start<end) {
        mergeCells(wb,sheet="splits",reference=paste(LETTERS[5*(i-1)+1],start,":",LETTERS[5*(i-1)+1],end,sep=""))
        mergeCells(wb,sheet="splits",reference=paste(LETTERS[5*(i-1)+2],start,":",LETTERS[5*(i-1)+2],end,sep=""))
        mergeCells(wb,sheet="splits",reference=paste(LETTERS[5*(i-1)+3],start,":",LETTERS[5*(i-1)+3],end,sep=""))
        mergeCells(wb,sheet="splits",reference=paste(LETTERS[5*(i-1)+4],start,":",LETTERS[5*(i-1)+4],end,sep=""))
        mergeCells(wb,sheet="splits",reference=paste(LETTERS[5*(i-1)+5],start,":",LETTERS[5*(i-1)+5],end,sep=""))
      }
      if(i>1) {
        writeWorksheet(wb,node,sheet="splits",startCol=5*(i-1)+1,startRow=start,header=FALSE)
        writeWorksheet(wb,paste(z$splitvar[z$node==node],z$operator[z$node==node],z$catlist[z$node==node],sep=""),sheet="splits",startCol=5*(i-1)+2,startRow=start,header=FALSE)
        data$logical=eval(parse(text=paste(paste("data[,'",z$splitvar[z$node %in% parent(node)],"']",z$operator[z$node %in% parent(node)],
                                                 z$catlist[z$node %in% parent(node)],sep="")[-1],collapse=" & ")))
        cluster=data[data$logical,]
        writeWorksheet(wb,length(unique(cluster[,id])),sheet="splits",startCol=5*(i-1)+3,startRow=start,header=FALSE)
        writeWorksheet(wb,sum(cluster[,denominator],na.rm=TRUE),sheet="splits",startCol=5*(i-1)+4,startRow=start,header=FALSE)
        writeWorksheet(wb,sum(cluster[,numerator],na.rm=TRUE)/sum(cluster[,denominator],na.rm=TRUE),sheet="splits",startCol=5*(i-1)+5,startRow=start,header=FALSE)  
      }
      if(i==1) {
        writeWorksheet(wb,node,sheet="splits",startCol=5*(i-1)+1,startRow=start,header=FALSE)
        writeWorksheet(wb,"All",sheet="splits",startCol=5*(i-1)+2,startRow=start,header=FALSE)
        writeWorksheet(wb,length(unique(data[,id])),sheet="splits",startCol=5*(i-1)+3,startRow=start,header=FALSE)
        writeWorksheet(wb,sum(data[,denominator],na.rm=TRUE),sheet="splits",startCol=5*(i-1)+4,startRow=start,header=FALSE)
        writeWorksheet(wb,sum(data[,numerator],na.rm=TRUE)/sum(data[,denominator],na.rm=TRUE),sheet="splits",startCol=5*(i-1)+5,startRow=start,header=FALSE)
      }
      cat("+ Node",node," ")
    }
    gc()
  }
  
  saveWorkbook(wb)
  
  wb <- loadWorkbook(paste("treeSummary_",tstamp,".xlsx",sep=""))
  createSheet(wb, name = "distribution")
  cat("\n\nPrinting distribution\n---------------------\n")
  
  z=z[order(as.numeric(z$node)),]
  nodes=as.numeric(z$node)
  for(i in 2:length(nodes)) {
    node=nodes[i]
    writeWorksheet(wb,t(c("Node",z[z$node==node,"splitvar"],"Count",denominator,"Y")),sheet="distribution",startCol=6*(i-1)+1,startRow=2,header=FALSE)
    writeWorksheet(wb,node,sheet="distribution",startCol=6*(i-1)+1,startRow=3,header=FALSE)
    data$logical=eval(parse(text=paste(paste("data[,'",z$splitvar[z$node %in% parent(node)],"']",z$operator[z$node %in% parent(node)],
                                             z$catlist[z$node %in% parent(node)],sep="")[-1],collapse=" & ")))
    cluster=data[data$logical,]
    summ0=aggregate(as.formula(paste(id,z[z$node==node,"splitvar"],sep="~")),cluster,FUN=function(w) length(unique(w)))
    summ=aggregate(as.formula(paste(numerator,z[z$node==node,"splitvar"],sep="~")),cluster,FUN=function(w) sum(as.numeric(w),na.rm=TRUE))
    summ2=aggregate(as.formula(paste(denominator,z[z$node==node,"splitvar"],sep="~")),cluster,FUN=function(w) sum(w,na.rm=TRUE))
    summ=merge(summ0,summ,by.x=colnames(summ0)[1],by.y=colnames(summ)[1],all.x=TRUE,all.y=TRUE)
    summ=merge(summ,summ2,by.x=colnames(summ)[1],by.y=colnames(summ2)[1],all.x=TRUE,all.y=TRUE)
    summ$metric=summ[,3]/summ[,4]
    writeWorksheet(wb,summ[,c(1,2,4,5)],sheet="distribution",startCol=6*(i-1)+2,startRow=3,header=FALSE)
    cat("+ Node",node," ")
    gc()
  }

  saveWorkbook(wb)
  
  cat("\n\nOutput file",paste("treeSummary_",tstamp,".xlsx",sep=""),"saved at location",getwd(),"\n")
  
}

