<template>
    <div class="scroll">
        <div class="container" @mousemove="mousemove">
            <div class="base"></div>
            <div class="left" ref="left">
                <img src="../../static/image/kda-karsa.jpg" alt="">
            </div>
        </div>
    </div>
</template>
<script>
export default {
    data() {
        return {
            posX: 0
        }
    },
    methods: {
        mousemove(e) {
            // 获取x 坐标
            this.posX = e.offsetX 
        }
    },
    watch: {
        posX(curX) {
            this.$refs.left.style.width = `${curX}px`
        }   
    }
}
</script>
<style lang="scss" scoped>
.scroll{
    .container{
        width: 960px;
        height: 540px;
        background-color: #cccccc;
        position: relative;
        cursor: ew-resize;
        .base{
            position: absolute;
            width: 960px;
            height: 540px;
            top: 0;
            left: 0;
            background: url('../../static/image/kda-karsa-golden.jpg') no-repeat;
            background-size: 100%;
        }
        .left{
            position: absolute;
            width: 480px;
            height: 540px;
            overflow: hidden;
            top: 0;
            left: 0;
            img{
              width: 960px;
              height: 540px;  
            }
        }
    }
}
</style>
