prepare_bat<-function(tmp_dir,primer,primer3dir,nprocess){
	cwd<-getwd();
	id<-c(paste(primer[,1],"_F",sep=""),paste(primer[,1],"_R",sep=""))
	seq<-c(primer[,2],primer[,3])
	x<-cbind(id,seq)

	idx1<-rep(1:length(seq),times=length(seq):1)
	idx2<-rep(1:length(seq),)

	cmd<-lapply(1:(nrow(x)-1),function(i){
		sapply((i+1):nrow(x),function(j){
			cmd<-paste("\"",primer3dir,"/ntthal.exe\" -s1 ",x[i,2]," -s2 ",x[j,2]," -dv 3.2 -n 0.85 -mv 52 -d 40 -a END1 -path \"",primer3dir,"/primer3_config/\"",sep="")
		})
	})
	cmd<-unlist(cmd)
	name<-lapply(1:(nrow(x)-1),function(i){
		sapply((i+1):nrow(x),function(j){
			paste(x[i,1],"#",x[j,1],sep="")
		})
	})
	name<-unlist(name)

	n<-nprocess
	idx<-rep(1:n,each=as.integer(length(cmd)/n)+1)
	idx<-idx[1:length(cmd)]
	cmd1<-split(cmd,idx)
	name1<-split(name,idx)

	xx<-mapply(function(x,y,i){
		y<-paste("echo '##",y,"'>> ",paste("\"",tmp_dir,"/part_",i,".txt\"",sep=""),sep="")
		x<-paste(x,">>",paste("\"",tmp_dir,"/part_",i,".txt\"",sep=""))
		cmd<-as.character(t(cbind(y,x)))
		writeLines(cmd,paste(tmp_dir,"/heterodimer_cmd_part_",i,".bat",sep=""))	
	},cmd1,name1,1:length(cmd1),SIMPLIFY=F)

	xx<-sapply(1:n,function(x){
		paste("\"",tmp_dir,"/heterodimer_cmd_part_",x,".bat\"",sep="")
	})
	writeLines(xx,paste(tmp_dir,"/batch_run.bat",sep=""))
	xx
}

output_result<-function(tmp_dir,primer,outputfile){
	cwd = getwd();
	setwd(tmp_dir)
	id<-c(paste(primer[,1],"_F",sep=""),paste(primer[,1],"_R",sep=""))
	seq<-c(primer[,2],primer[,3])
	x<-cbind(id,seq)
	file<-list.files(pattern="*part_\\d+.txt") 
	dat<-lapply(file,readLines)
	names(dat)<-substr(file,6,nchar(file)-4)
	dat<-dat[order(as.integer(names(dat)))]
	dat<-unlist(dat)
	writeLines(dat,"all.txt")

	fr<-grep("##",dat)
	id<-dat[fr]
	id<-substr(id,4,nchar(id)-1)
	deltaG<-dat[fr+1]
	deltaG<-substr(deltaG,regexpr("dG = ",deltaG)+4,regexpr("t =",deltaG)-1)
	product.size<-nchar(dat[fr+2])-4
	match.len<-sapply(gregexpr("[ACGT]",dat[fr+3],perl=T),length)

	group.idx<-cut(1:length(dat),breaks=c(fr,length(dat)),include.lowest=T,right=F)
	# dat.split<-split(dat,group.idx)
	# library(parallel)
	# cl<-makeCluster(20)
	match.pattern<-tapply(dat,group.idx,function(x){
		if(length(x)==6){
			seq<-do.call(rbind,strsplit(x[3:6],""))[,-c(1:4)]
			idx1<-seq[2,]%in%c("A","C","G","T")
			idx2<-seq[3,]%in%c("A","C","G","T")
			seq[1,][idx1]<-seq[2,][idx1]
			seq[4,][idx2]<-seq[3,][idx2]
			seq[2,][idx1]<-"|"
			paste(apply(seq[-3,],1,paste,collapse=""),collapse="\n")
		}else{
			""
		}
	})
	#stopCluster(cl)

	dat1<-data.frame(ID=id,DeltaG=as.numeric(deltaG),MatchPattern=match.pattern,ProductSize=product.size,MatchLenth=match.len,stringsAsFactors=F)

	setwd(tmp_dir)
	dat1<-dat1[order(dat1$DeltaG),]

	seq<-strsplit(dat1[,3],"\n",fixed=T)
	end.ol<-t(sapply(seq,function(x){
		base.idx<-gregexpr("A|C|G|T",x[c(1,3)])
		match.idx<-gregexpr("|",x[2],fixed=T)[[1]]
		if(!match.idx[1]%in%NA){
			a.idx<-which(base.idx[[1]]%in%match.idx)-length(base.idx[[1]])
			b.idx<- -(which(base.idx[[2]]%in%match.idx)-1)
			end.shift<- -(max(a.idx)+max(b.idx))
			gap.and.mismatch<- sum(!min(match.idx):max(match.idx)%in%match.idx)
			c(length(match.idx)-3*end.shift-2^gap.and.mismatch,end.shift,gap.and.mismatch)
		}else{
			c(-1,-1,-1)
		}
	}))

	colnames(end.ol)<-c("Score","End.Shift","Gap.and.Mismatch")
	dat2<-cbind(dat1,end.ol)
	#dat2<-dat2[order(dat2[,2]*end.ol2),]

	primer.seq<-x[,2]
	primer.id<-x[,1]

	id<-do.call(rbind,strsplit(dat2$ID,"#"))
	dat2<-cbind(dat2,primer.seq[match(id[,1],primer.id)],stringsAsFactors=F)
	dat2<-cbind(dat2,primer.seq[match(id[,2],primer.id)],stringsAsFactors=F)
	dat2<-cbind(id,dat2,stringsAsFactors=F)
	dat2<-dat2[,c(1,11,2,12,4,5,6,7,8,9,10)]
	colnames(dat2)[1:5]<-c("Primer1","Primer1.Seq","Primer2","Primer2.Seq","DeltaG (kcal/mol)")
	dat2[,5]<-round(dat2[,5]/1000,2)
	dat2<-dat2[order(dat2[,5]),]

	#dat2<-dat2[dat2[,5]+2*dat2[,10]+1.5*dat2[,11]< -6,]
	library(xlsx)
	setwd(cwd);
	write.xlsx(dat2,file=outputfile,row.names=F)
}
